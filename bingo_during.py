import os
from bingo_prep import calcBingoStats, openCalloutFile
from tkinter import filedialog
import webbrowser
import msvcrt

# Callout mode
def showCalloutsMade(mode=""):
    percentage = len(callouts_made) / len(randomised_item_bank) * 100
    print(f"\n{percentage:0.1f}% of callouts made ({len(callouts_made)}/{len(randomised_item_bank)}). ", end="")
    if len(callouts_made) == 1:
        print(f"Callout was '{callouts_made[0]}'.")
    else:
        print(f"Callouts were ", end="")
        for callout in callouts_made:
            if callout == callouts_made[-1] and mode != "?":
                print(f"and '{callout}'.\n")
            elif callout == callouts_made[-1] and mode == "?":
                print(f"and '{callout}'.", end =" ")
            else:
                print(f"'{callout}'", end=", ")

output_file = ""
while not output_file:
    output_file = filedialog.askopenfilename(title="Open existing callout sheet", filetypes=[("Text file", "*.txt")])

callout_file, _ = os.path.splitext(output_file)
callout_file = f"{callout_file}.txt"

randomised_item_bank = openCalloutFile(callout_file)
calcBingoStats(randomised_item_bank)

choice = input("Press ENTER to start callouts.")
if choice == "":
    choice = input("Press ENTER to automatically show Cambridge dictionary entry.")
    
    if choice == "":
        autoGoToCambridge = True
    else:
        autoGoToCambridge = False

    print("Press ENTER for next callout, ? for callouts made, and any other key to stop.\n")

    callingOut = True
    i = 0
    callouts_made = []
    while callingOut:
        callout = randomised_item_bank[i]
        url = f"https://dictionary.cambridge.org/dictionary/english/{callout}"

        if autoGoToCambridge:
            webbrowser.open(url, new=2, autoraise=True)

        choice = input(callout)
        callouts_made.append(callout)

        i += 1
        if i > len(randomised_item_bank) - 1:       #Stop if end
            showCalloutsMade()
            callingOut = False
            print("That was the last callout...\n")
            msvcrt.getch()
        elif choice == "?":                         #Show callouts
            showCalloutsMade(mode="?")
            print("Press any key to continue calling out...\n")
            msvcrt.getch()
        elif choice != "":                          #Stop if not ENTER key
            showCalloutsMade()
            callingOut = False