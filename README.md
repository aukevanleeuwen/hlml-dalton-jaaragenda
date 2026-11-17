# hlml-dalton-jaaragenda

The Dalton location of the Haarlemmermeer Lyceum publishes its *jaaragenda*
(school year calendar) as an Excel file. This project turns the columns you care
about (for example *Algemeen* and your child's *leerjaar*) into calendar files
(`.ics`) you can subscribe to in Google Calendar, Apple Calendar or Outlook.
Each column becomes its own calendar, so each child gets its own color.

The calendars of this repository are in [`calendars/`](calendars/). They are
kept up to date automatically (see [Updates](#updates)).

## Subscribing in Google Calendar

This only works in the browser, not in the phone app. Do it once per `.ics`
file:

1. Open [calendar.google.com](https://calendar.google.com).
2. Click **+** next to **Other calendars** and choose **From URL**.
3. Paste the link to the raw file, for example:
   `https://raw.githubusercontent.com/aukevanleeuwen/hlml-dalton-jaaragenda/main/calendars/immy.ics`
4. Click **Add calendar**.
5. Hover over the new calendar in the list, click the three dots and pick a
   color.

The calendar then also shows up in the phone app. Google polls for changes on
its own schedule, which can take up to a day.

**Apple Calendar (iPhone):** Settings → Calendar → Accounts → Add Account →
Other → Add Subscribed Calendar, and paste the same link.

**Rather not subscribe?** Download the `.ics` file and import it into an empty
calendar (Google Calendar: Settings → Import & export). You will have to repeat
that for every new version of the jaaragenda.

## What ends up in the calendar

- Every line in a cell becomes an all-day event, prefixed with the name:
  `Sam: Schoolfotograaf`. Events are marked as "free", so they do not block
  your own calendar.
- The same text on consecutive days becomes one multi-day event.
- Vacations ("Start kerstvakantie" up to "Laatste dag kerstvakantie") become one
  block.
- An event that applies to two children shows up in both calendars.
- Event IDs are derived from the content, so a new version of the jaaragenda
  updates subscribed calendars without duplicating events.

## Updates

The [Update calendars](.github/workflows/calendars.yml) workflow runs every
morning. It fetches the jaaragenda via
[haarlemmermeerlyceum.nl/jaaragenda-dalton](https://haarlemmermeerlyceum.nl/jaaragenda-dalton),
a stable link that redirects to the current spreadsheet. When there is a new
version it replaces the one in [`source/`](source/) (older versions stay in the
git history), regenerates the `.ics` files and commits both. When nothing
changed, it does nothing.

To run it right away: **Actions** → **Update calendars** → **Run workflow**.

The workflow also runs on a push that changes `source/`, `columns.toml` or
`convert.py`; it then converts the most recently added spreadsheet in
`source/`. Note that the next daily run replaces a manually added spreadsheet
with the one from the website if they differ.

GitHub disables scheduled workflows in public repositories after 60 days
without activity (the summer holiday, for instance). You get an email when that
happens; re-enable it from the **Actions** tab.

## Using it for your own children

1. Fork this repository.
2. Delete `columns.toml` and the files in `calendars/`.
3. Install [uv](https://docs.astral.sh/uv/) and run:

   ```bash
   uv run convert.py --fetch
   ```

   The script downloads the newest jaaragenda, lists its columns and asks which
   ones you want and what name to use for each. Your choice is saved to
   `columns.toml`; edit or delete that file to choose again.
4. Commit and push. In a fork, enable workflows from the **Actions** tab, or
   new versions will not be picked up automatically.
5. Subscribe as described above, with your GitHub user name and file names in
   the link.

`columns.toml` looks like this:

```toml
[columns]
"Algemeen" = "Dalton"
"Leerjaar 1" = "Sam"
"Leerjaar 3VWO" = "Noor"
```

### Options

```text
uv run convert.py [XLSX | --fetch [URL]] [--column "COLUMN=NAME" ...] [--config FILE] [--output DIR]
```

- `XLSX`: the spreadsheet to convert. Defaults to the only `.xlsx` file in
  `source/`.
- `--fetch`: first download the newest jaaragenda from the school website into
  `source/`, replacing the old one. Pass a URL to download it from somewhere
  else.
- `--column`: pick a column and name without `columns.toml`, for example
  `--column "Leerjaar 2=Sam"`. Repeatable.
- `--config`: use another file than `columns.toml`.
- `--output`: write to another directory than `calendars/`.

The script expects column headers in the first row, including a `Datum`
column. Every other column except `Week` and `Dag` can be chosen.
