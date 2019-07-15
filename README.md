# O365 Connectivity

Import your Office 365 calendar into Google Calendar and more.

## Getting Started

### Prerequisites

TODO

### Setup

1. Configure O365 by following the steps in [O365: Authentication Flow](https://github.com/O365/python-o365#authentication-flow).
2. `./o365cal setup` and specify your client_id & client_secret
3. `./o365cal login` to log into your Office 365 account

## Usage

```bash
./o365cal generate > MyCal.ics
```

## Play with Others

### Import using gcalcli

```bash
pip install gcalcli vobject
```

```bash
gcalcli --calendar 'My Outlook Cal' import MyCal.ics
```

### Gist-based Approach

We can upload MyCal.ics to <https://gist.github.com/> manually or via CLI using [gist CLI](https://github.com/defunkt/gist).

1. Go to <https://github.com/settings/tokens> and press "Generate New Token". Copy the token to `~/.gist`.
2. Create a secret gist via `gist -p MyCal.isc` and note the URL which also includes the GIST_ID.
3. In Google Calendar,[import a calendar from URL](https://calendar.google.com/calendar/r/settings/addbyurl) and specify the URL + '/raw'

You can then update via...

```bash
./o365cal generate > MyCal.ics
gist -u {URL or GIST_ID} MyCal.ics
```

## Built With

- [python-o365](https://github.com/O365/python-o365) - interact with Microsoft Graph and Office 365 API
- [click](https://github.com/pallets/click/) - CLI toolkit
- [openpyxl](https://github.com/collective/icalendar) - RFC 5545 compatible parser/generator for iCalender files

