# merge-outlook-contacts

This small ruby script helps to get rid of duplicated contacts from [outlook-people](https://outlook.live.com/owa/?path=/people). I had at one day around 2500 contacts listed, 5 times more than expected and most of them were exact duplicates.
The script merges contacts with the same first- and last-name together. When a merge-conflicts appear, the script tries to find an empty field with asimilar name, and fills in the conflicted entry there.

## How it works
- Export (download) all your contacts at [outlook-people](https://outlook.live.com/owa/?path=/people) as a .csv-file in the `Microsoft Outlook-CSV` format.
- run the script with `ruby .rb path/to/exported/contacts.csv`.This results in a new file in `path/to/exported/contacts_merged.csv`, which contains only  contacts with a unique name.
- Delete all contacts on outlook (Be sure to go `Your Contacts>Contacts`, otherwise no delete-option will be shown. Tip: At the left of the title 'Contacts' appears a hidden checkbox on hover over, which selects all your contacts at one glance)
- Import the merged contact-csv..
- Enjoy the uniqueness of your friends ;)

## Caveats
- The script supposes exactly 61 entries per line. When you have for some reasons commas inside a field (e.g. `First Name: Reto, Holz`), it is not possible to decide wheter `Holz` is related to the first- or the middle-name! When too many fields are present, then the surplus fields are concatenated with last field (notes). When less than `61` entries are present, the contact is skipped.
- When no name-field is present, the entry for this contact will be skipped. 
