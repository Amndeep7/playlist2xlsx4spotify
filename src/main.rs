use futures_util::{StreamExt, TryStreamExt};
use itertools::Itertools;
use rspotify::{
    AuthCodePkceSpotify, Credentials, OAuth,
    clients::{BaseClient, OAuthClient},
    model::{PlayableItem, playlist::SimplifiedPlaylist},
    scopes,
};
use rust_xlsxwriter::{Color, Format, Table, TableColumn, TableStyle, Workbook};
use sanitise_file_name::sanitise;
use serde::Deserialize;
use std::ffi::OsString;
use std::future;
use std::io;
use std::path::PathBuf;
use std::vec;

async fn get_spotify_client() -> AuthCodePkceSpotify {
    let creds = Credentials::from_env().unwrap();
    let oauth = OAuth::from_env(scopes!("playlist-read-private")).unwrap();
    let mut spotify = AuthCodePkceSpotify::new(creds, oauth);
    let url = spotify.get_authorize_url(None).unwrap();
    spotify.prompt_for_token(&url).await.unwrap();
    spotify
}

// spotify has restricted the APIs considerably to prevent AI companies from scraping all their
// data, now, while we can list all playlists, we can only read the tracklist for playlists that we
// own or are collaborators on
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

// dunno if it's just not getting paged properly or another restriction spotify implemented was
// capping out at 50 songs per playlist
async fn get_playlist_tracks_data(
    client: &AuthCodePkceSpotify,
    playlist: &SimplifiedPlaylist,
) -> Vec<LimitedTrackData> {
    // href,limit,offset,total,items.is_local seem to be required by rspotify probably for iteration
    // purposes since it worked just fine without them on the spotify api site
    let stream = client.playlist_items(
        playlist.id.clone(),
        Some("href,limit,offset,total,items(is_local,item(name,album(name),artists(name)))"),
        None,
    );
    let raw_data = stream
        .collect::<Vec<_>>()
        .await
        .into_iter()
        .collect::<Result<Vec<_>, _>>()
        .unwrap();
    let data: Vec<LimitedTrackData> = raw_data
        .into_iter()
        .map(|track_data| match track_data.item.unwrap() {
            PlayableItem::Unknown(value) => {
                serde_json::from_value::<LimitedTrackData>(value.clone()).unwrap()
            }
            _ => panic!("Spotify returned something unexpected"),
        })
        .collect();
    data
}

fn generate_filename(playlist: &SimplifiedPlaylist) -> OsString {
    let sanitized = sanitise(&(playlist.name.clone() + ".xlsx"));
    let path = PathBuf::from(sanitized);
    let basename = path.as_path().file_name().unwrap().to_os_string();
    println!(
        "Using filename {:?} for playlist \"{}\"",
        basename, playlist.name
    );
    basename
}

impl IntoIterator for LimitedTrackData {
    type Item = String;
    type IntoIter = vec::IntoIter<Self::Item>;

    fn into_iter(self) -> Self::IntoIter {
        vec![
            self.name,
            self.album.name,
            self.artists
                .into_iter()
                .map(|artist| artist.name)
                .collect::<Vec<_>>()
                .join(";"),
        ]
        .into_iter()
    }
}

// uses the interquartile range algorithm to remove outliers and then returns the number of
// characters in the remaining longest string
fn generate_column_widths(data: &[LimitedTrackData]) -> Vec<usize> {
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
    let mut widths = all_lengths
        .iter_mut()
        .map(|lengths| {
            lengths.sort();
            let q1 = lengths[lengths.len() / 4];
            let q3 = lengths[3 * lengths.len() / 4];
            let interquartile_range = q3 - q1;
            lengths.retain(|&length| 2 * length <= 2 * q3 + 3 * interquartile_range); // mult by 2 so that I don't need to deal with floats
            lengths.iter().copied().max().unwrap().min(30)
        })
        .collect::<Vec<_>>();
    widths.push(30); // the 'vote' tab also needs a defined width
    widths
}

fn generate_xlsx(data: Vec<LimitedTrackData>) -> Workbook {
    let mut workbook = Workbook::new();
    let format = Format::new()
        .set_background_color(Color::Black)
        .set_font_color(Color::White)
        .set_font_name("Aptos Narrow")
        .set_font_size(11);
    workbook.set_default_format(&format, 20, 64).unwrap();
    let columns = vec!["Track Name", "Album Name", "Artist Name(s)", "Vote"]
        .into_iter()
        .map(|header| TableColumn::new().set_header(header))
        .collect::<Vec<_>>();
    let table = Table::new()
        .set_columns(&columns)
        .set_style(TableStyle::Dark1);
    let worksheet = workbook.add_worksheet();
    worksheet
        .add_table(0, 0, data.len().try_into().unwrap(), 3, &table)
        .unwrap();
    generate_column_widths(&data)
        .iter()
        .enumerate()
        .for_each(|(i, &width)| {
            worksheet.set_column_width(i as u16, width as f64).unwrap();
        });
    worksheet.write_row_matrix(1, 0, data).unwrap();
    workbook
}

#[tokio::main]
async fn main() {
    let client = get_spotify_client().await;
    let playlists = get_permitted_playlists(&client).await;
    let playlist = select_playlist(&playlists);
    let tracks_data = get_playlist_tracks_data(&client, playlist).await;
    let filename = generate_filename(playlist);
    let mut workbook = generate_xlsx(tracks_data);
    workbook.save(filename).unwrap();
}
