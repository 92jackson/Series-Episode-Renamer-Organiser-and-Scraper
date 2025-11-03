# Episode Datasheets

This folder contains CSV episode data used by Episode Organiser.

- `thomas_&_friends_(1984).csv`: episode data based on https://en.wikipedia.org/wiki/List_of_Thomas_%26_Friends_episodes, movie entries are based on https://www.imdb.com/list/ls098617186/.
- `jack_&_the_sodor_construction_company_(2006).csv`: episode data based on https://www.imdb.com/title/tt2426370/episodes/

## Expected CSV Format

- Header row is required: `ep_no,series_ep_code,title,air_date`
- Rows are comma‑separated, typically quoted (RFC‑4180 style):

```
"001","s01e01","Thomas & Gordon","1984-10-09"
"321.5","s00e04","Hero of the Rails","2009-09-08"
```

- Column meanings:
- `ep_no`: overall episode number. Use decimals for films/specials if needed (e.g. `321.5`).
- `series_ep_code`: `sXXeXX` for episodes, `s00eXX` for films/specials.
- `title`: episode or film title.
- `air_date`: ISO date `YYYY-MM-DD`. Leave blank if unknown.

- Encoding: UTF‑8 is recommended.
- Extra columns are ignored.
- Optional columns recognised for alternate title matching:
- `alt_title`, `alt_titles` (semicolon‑separated), `aka`, `alternate_title`.
- These help match files that use different wording for the episode title.

### Filename character sanitization

- Titles may include characters illegal on Windows (e.g. `?:<>"/\|*:`).
- During renaming, these are automatically sanitized:
  - `:` becomes `-`, `/` and `\` become `-`.
  - `?`, `*`, `<`, `>`, `"` are removed.
  - Excess spaces are collapsed and trimmed.
- You do not need to manually edit titles in the CSV to remove these.

### Example with an alternate title

Header:

```
"ep_no","series_ep_code","title","air_date","alt_title"
```

Row:

```
"013","s01e13","Percy Helps Out","2006-10-02","Jack & Jet Pack - Percy Helps Out"
"528.5","m14","Big World! Big Adventures!","2018-11-12"
```

- Season mapping: episodes go to `Season N` from `sXX`, movies go to `Movies/Movie Name (Year)/`.
