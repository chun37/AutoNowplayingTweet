import win32com.client
import tweepy
import keys
import os
import time
from enum import Enum


class ArtworkFormat(Enum):
    JPG = 1
    PNG = 2
    BMP = 3


class NowPlayingTweet():
    def __init__(self, itunes, twitter, interval=5):
        self.twitter = twitter
        self.itunes = itunes
        self.last_track = None
        if itunes.CurrentTrack:
            self.last_track = itunes.CurrentTrack.GetITObjectIDs()
        self.interval = interval
        self.tmpfile = os.getcwd() + "\\tmp."

    def tweet(self, title, artist, image):
        if image:
            media_id = self.twitter.media_upload(image).media_id
            os.remove(image)
        else:
            media_id = ""
        text = f"#NowPlaying {title} - {artist}"
        self.twitter.update_status(text, media_ids=[media_id])

    def fetchTrack(self):
        track = self.itunes.CurrentTrack
        if not track or self.itunes.PlayerState == 0:
            return
        track_id = track.GetITObjectIDs()
        if track_id == self.last_track:
            return
        self.last_track = track_id
        artist = track.Artist
        title = track.Name
        artwork = track.Artwork.Item(1)
        if artwork:
            ext = ArtworkFormat(artwork.Format).name
            image_path = self.tmpfile + ext
            artwork.SaveArtworkToFile(image_path)
            artwork = image_path
        else:
            artwork = None
        self.tweet(title, artist, artwork)
        time.sleep(self.interval)


if __name__ == "__main__":
    itunes = win32com.client.gencache.EnsureDispatch("iTunes.Application")
    consumer_key = os.environ["consumer_key"]
    consumer_secret = os.environ["consumer_secret"]
    access_token = os.environ["access_token"]
    access_token_secret = os.environ["access_token_secret"]
    auth = tweepy.OAuthHandler(consumer_key, consumer_secret)
    auth.set_access_token(access_token, access_token_secret)
    twitter = tweepy.API(auth)
    np = NowPlayingTweet(itunes, twitter)
    while True:
        np.fetchTrack()
