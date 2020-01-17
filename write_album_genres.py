from xlrd import open_workbook
from xlwt import Workbook
from xlutils.copy import copy as xlcopy # http://pypi.python.org/pypi/xlutils
from functools import reduce

from copy import copy

from xml.etree import ElementTree

import requests
from requests_oauthlib import OAuth1

from keys import consumer_key, consumer_secret, auth_token, auth_token_secret, spreadsheet_loc

output_file = open("./output.txt", "w+")

def get_auth_token():
    URL = "http://api.music-story.com/oauth/request_token"

    auth = OAuth1(consumer_key, consumer_secret)

    resp = requests.get(URL, auth=auth)

    print(resp.status_code)
    print(resp.text)

    tree = ElementTree.fromstring(resp.content)
    token = None
    token_secret = None
    for child in tree.iter('*'):
        if child.tag == "token":
            token = child.text
        elif child.tag == "token_secret":
            token_secret = child.text

    assert token is not None
    assert token_secret is not None

    return token, token_secret


def get_id_from_album(album, artist):
    URL = "http://api.music-story.com/album/search"
    PARAMS = {
        "title": album
    }

    auth = OAuth1(consumer_key, consumer_secret, auth_token, auth_token_secret)

    resp = requests.get(url=URL, params=PARAMS, auth=auth)

    print(resp.status_code)
    assert resp.status_code == 200

    try:
        output_file.write(str(resp.text) + "\n")
    except:
        print("unable to print", resp.text)

    tree = ElementTree.fromstring(resp.content)
    results = []
    curr = {}
    next_is_score = False
    for child in tree.iter('*'):
        if next_is_score:
            curr["score"] = float(child.text)
            results.append(copy(curr))
            curr = {}
            next_is_score = False
            continue

        if child.tag == "id":
            curr["id"] = int(child.text)
        elif child.tag == "search_scores":
            next_is_score = True
        elif child.tag == "url":
            curr["url"] = child.text

    if not len(results):
        return None

    max_score = max(i["score"] for i in results)
    best_matches = [x for x in results if x["score"] == max_score]

    for x in best_matches:
        word_matches = 0
        for word in artist.split():
            if word.lower() in x["url"].lower():
                word_matches += 1
        x["word_matches"] = word_matches

    most_word_matches = max(i["word_matches"] for i in best_matches)
    id = next(x for x in best_matches if x["word_matches"] == most_word_matches)["id"]

    return id

def get_genres_from_id(id):
    URL = "http://api.music-story.com/en/album/" + str(id) + "/genres"
    auth = OAuth1(consumer_key, consumer_secret, auth_token, auth_token_secret)
    resp = requests.get(url=URL, auth=auth)

    print(resp.status_code)
    assert resp.status_code == 200
    print(resp.text)

    tree = ElementTree.fromstring(resp.content)
    genres = []
    for child in tree.iter('*'):
        if child.tag == "name":
            genres.append(child.text)

    output_file.write("genres: " + str(genres) + "\n")

    return genres

r_book = open_workbook(spreadsheet_loc, formatting_info=False)
r_sheet = r_book.sheet_by_index(0)
w_book = xlcopy(r_book) # cant write and read to same excel file unfortunately
w_sheet = w_book.get_sheet(0)

if not auth_token or not auth_token_secret:
    print("getting auth token")
    auth_token, auth_token_secret = get_auth_token()


all_genres = set()

r = 3

genre_col = 8

# go through r_sheet and make api call for each
for row in r_sheet.get_rows():
    album = str(row[1].value)
    artist = str(row[2].value)

    if not album or not artist or (album == "Album" and artist == "Artist"):
        continue

    if row[genre_col].value:
        print("skipping", album, artist)

        for genre in row[genre_col].value.split(","):
            all_genres.add(genre.strip())

        r += 1
        continue

    print(str(r - 3), "get", album, artist)
    try:
        output_file.write("get " + album + " " + artist + '\n')
    except:
        output_file.write("couldnt write name")

    album_id = get_id_from_album(album, artist)

    if album_id == None:
        output_file.write("genres: " + str(["COULD NOT GET"]) + "\n")
        genres = ["COULD NOT GET"]
    else:
        genres = get_genres_from_id(album_id)

    print(genres)
    for genre in genres:
        all_genres.add(genre)

    # write to spreadsheet

    if len(genres):
        genre_str = reduce(lambda g, p: g + ", " + p, genres)
    else:
        genre_str = "COULD NOT GET"

    w_sheet.write(r, genre_col, genre_str)

    w_book.save("Albums-out.xls")

    r += 1

    # if r >= 5:
    #     break

print("all genres", all_genres)
genre_list_col = 20
r = 2
for genre in all_genres:
    w_sheet.write(r, genre_list_col, genre)
    r += 1

w_book.save("Albums-out.xls")
