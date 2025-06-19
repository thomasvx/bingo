from bingo_prep import jochems_quiz, laatsteA, laatsteB
from bingo_prep import report_partial_duplicates, remove_full_duplicates
import random

s = [jochems_quiz, laatsteA, laatsteB]
bigList = []
for setje in s:
    idx = 0
    songs = []
    artists = []
    for songOrArtist in setje:
        if idx % 2 == 0:
            songs.append(songOrArtist)
        else:
            artists.append(songOrArtist)
        idx += 1

    combos = []
    for i in range(len(artists)):
        track = f"{artists[i]} - {songs[i]}"
        combos.append(track)
        # bigList.append(songs[i])

    # random.shuffle(combos)
    for combo in combos:
        print(f"{combo}")
    print(len(combos))

# bigList = remove_full_duplicates(bigList)
# report_partial_duplicates(bigList)

#Git test