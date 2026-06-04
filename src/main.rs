use std::{
    error::Error,
    ffi::OsString,
    fs::{File, OpenOptions},
    future,
    io,
    num::TryFromIntError,
    path::{Path, PathBuf},
    vec,
};

use futures_util::TryStreamExt;
use itertools::Itertools;
use rspotify::{
    AuthCodePkceSpotify,
    Credentials,
    OAuth,
    clients::{BaseClient, OAuthClient},
    model::{PlayableItem, playlist::SimplifiedPlaylist},
    scopes,
};
use rust_xlsxwriter::{Color, Format, Table, TableColumn, TableStyle, Workbook};
use sanitise_file_name::sanitise;
use serde::Deserialize;

#[derive(Debug, Deserialize)]
struct LimitedAlbumData {
    name: String,
}

#[derive(Debug, Deserialize)]
struct LimitedArtistData {
    name: String,
}

#[derive(Debug, Deserialize)]
struct LimitedTrackData {
    name: String,
    album: LimitedAlbumData,
    artists: Vec<LimitedArtistData>,
}

impl IntoIterator for LimitedTrackData {
    type IntoIter = vec::IntoIter<Self::Item>;
    type Item = String;

    fn into_iter(self) -> Self::IntoIter {
        vec![
            self.name,
            self.album.name,
            self.artists.into_iter().map(|artist| artist.name).join(";"),
        ]
        .into_iter()
    }
}

async fn get_spotify_client() -> AuthCodePkceSpotify {
    let creds = Credentials::from_env().unwrap();
    let oauth = OAuth::from_env(scopes!("playlist-read-private")).unwrap();
    let mut spotify = AuthCodePkceSpotify::new(creds, oauth);
    let url = spotify.get_authorize_url(None).unwrap();
    spotify.prompt_for_token(&url).await.unwrap();
    spotify
}

// spotify has restricted the APIs considerably to prevent AI companies from scraping all their data, now, while we can list all playlists, we can only read the tracklist for playlists that we own or are collaborators on
async fn get_permitted_playlists(client: &AuthCodePkceSpotify) -> Vec<SimplifiedPlaylist> {
    let me = client.current_user().await.unwrap();
    client
        .current_user_playlists()
        .try_filter(|p| future::ready(p.owner.id == me.id || p.collaborative))
        .try_collect()
        .await
        .unwrap()
}

fn select_playlist(playlists: &[SimplifiedPlaylist]) -> &SimplifiedPlaylist {
    for (i, playlist) in playlists.iter().enumerate() {
        println!("{i}) {}", playlist.name);
    }
    println!(
        "Which playlist to process (NOTE: it must be owned or collaborated by you so you might need to manually make a copy if it doesn't show up in the list)?  Provide the index:"
    );
    let mut selection = String::new();
    io::stdin().read_line(&mut selection).unwrap();
    let selection: usize = selection.trim().parse().unwrap();
    println!("You chose \"{}\"", playlists[selection].name);
    &playlists[selection]
}

// dunno if it's just not getting paged properly or another restriction spotify implemented was capping out at 50 songs per playlist
async fn get_playlist_tracks_data(
    client: &AuthCodePkceSpotify,
    playlist: &SimplifiedPlaylist,
) -> Vec<LimitedTrackData> {
    // href,limit,offset,total,items.is_local seem to be required by rspotify probably for iteration purposes since it worked just fine without them on the spotify api site
    client
        .playlist_items(
            playlist.id.clone(),
            Some("href,limit,offset,total,items(is_local,item(name,album(name),artists(name)))"),
            None,
        )
        .map_ok(|track_data| match track_data.item {
            Some(PlayableItem::Unknown(value)) => {
                serde_json::from_value::<LimitedTrackData>(value).unwrap()
            },
            _ => panic!("Spotify returned something unexpected"),
        })
        .try_collect::<Vec<_>>()
        .await
        .unwrap()
}

// uses the interquartile range algorithm to remove outliers and then returns the number of characters in the remaining longest string
fn generate_column_widths(data: &[LimitedTrackData]) -> Result<Vec<u8>, TryFromIntError> {
    let (track_name_lengths, album_name_lengths, artist_name_lengths): (Vec<_>, Vec<_>, Vec<_>) =
        data.iter()
            .map(|track| {
                let artist_names_length = track
                    .artists
                    .iter()
                    .map(|artist| artist.name.len())
                    .sum::<usize>()
                    + (track.artists.len() - 1); // artist names are separated by semicolons (see trait impl)
                (
                    track.name.len(),
                    track.album.name.len(),
                    artist_names_length,
                )
            })
            .multiunzip();
    let mut all_lengths = [track_name_lengths, album_name_lengths, artist_name_lengths];
    let max_width = 30;
    let mut widths = all_lengths
        .iter_mut()
        .map(|lengths| -> Result<u8, TryFromIntError> {
            lengths.sort();
            let q1 = lengths[lengths.len() / 4];
            let q3 = lengths[3 * lengths.len() / 4];
            let interquartile_range = q3 - q1;
            let width = lengths
                .iter()
                .filter(|&length| 2 * length <= 2 * q3 + 3 * interquartile_range) // mult by 2 so that I don't need to deal with floats
                .max()
                .map_or(max_width, |&x| x)
                .min(max_width)
                .try_into()?; // guaranteed to work so long as max_width remains under the max size of a u8
            Ok(width)
        })
        .collect::<Result<Vec<_>, TryFromIntError>>()?;
    widths.push(max_width.try_into()?); // the 'vote' tab also needs a defined width
    Ok(widths)
}

fn generate_xlsx(data: impl Into<Vec<LimitedTrackData>>) -> Result<Workbook, Box<dyn Error>> {
    let data = data.into();
    let mut workbook = Workbook::new();
    let format = Format::new()
        .set_background_color(Color::Black)
        .set_font_color(Color::White)
        .set_font_name("Aptos Narrow")
        .set_font_size(11);
    workbook.set_default_format(&format, 20, 64)?;
    let columns = ["Track Name", "Album Name", "Artist Name(s)", "Vote"]
        .iter()
        .map(|&header| TableColumn::new().set_header(header))
        .collect::<Vec<_>>();
    let table = Table::new()
        .set_columns(&columns)
        .set_style(TableStyle::Dark1);
    let worksheet = workbook.add_worksheet();
    worksheet.add_table(0, 0, data.len().try_into()?, 3, &table)?;
    for (i, &width) in generate_column_widths(&data)?.iter().enumerate() {
        worksheet.set_column_width(i.try_into()?, f64::from(width))?;
    }
    worksheet.write_row_matrix(1, 0, data)?;
    Ok(workbook)
}

fn generate_filename(playlist: &SimplifiedPlaylist) -> Result<OsString, io::Error> {
    let suffix = "xlsx";
    let sanitized = sanitise(&format!("{}.{}", playlist.name, suffix));
    let sanitized = if sanitized == ".xlsx" {
        // excel on windows doesn't open dotfiles for some reason
        format!("_.{}", suffix)
    } else {
        sanitized
    };
    let path = PathBuf::from(sanitized);
    let basename = path
        .as_path()
        .file_name()
        .ok_or_else(|| {
            io::Error::new(
                io::ErrorKind::InvalidInput,
                "generated filename somehow has no filename",
            )
        })?
        .to_os_string();
    println!(
        "Using filename {:?} for playlist \"{}\"",
        basename, playlist.name
    );
    Ok(basename)
}

fn get_file_handle(path: impl AsRef<Path>) -> io::Result<File> {
    OpenOptions::new().write(true).create_new(true).open(path)
}

#[tokio::main]
async fn main() -> Result<(), Box<dyn Error>> {
    let client = get_spotify_client().await;
    let playlists = get_permitted_playlists(&client).await;
    let playlist = select_playlist(&playlists);
    let tracks_data = get_playlist_tracks_data(&client, playlist).await;
    let mut workbook = generate_xlsx(tracks_data)?;
    let filename = generate_filename(playlist)?;
    let handle = get_file_handle(&filename)?;
    workbook.save_to_writer(handle)?;
    Ok(())
}
