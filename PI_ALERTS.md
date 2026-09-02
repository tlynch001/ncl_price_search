# Raspberry Pi + Fastmail monitoring

This adds a Pi-oriented monitor without removing the existing Windows script.

## Why the alert logic changed

A single hard threshold such as `$320` can miss useful events. The monitor now
alerts when a sailing has a meaningful *event*:

- crosses below an absolute threshold (default `$350`)
- drops at least 10% from the immediately previous observation
- reaches a new 30-day low
- reaches a new all-time low

Repeated checks at the same cheap price do **not** keep emailing you. The next
alert requires a new event, such as a further price drop.

The existing CSV schema is preserved so previous history can be reused.

## Fastmail setup

Create a Fastmail app password for this monitor. Then on the Pi:

```bash
mkdir -p ~/.config
cp fastmail.env.example ~/.config/ncl-price-search.env
nano ~/.config/ncl-price-search.env
chmod 600 ~/.config/ncl-price-search.env
```

`NCL_SMTP_PASSWORD` must be the Fastmail app password, not your normal account
password. The monitor uses `smtp.fastmail.com:587` with TLS.

## Test manually first

From the repo directory:

```bash
set -a
source ~/.config/ncl-price-search.env
set +a
pwsh -NoProfile -File ./Monitor-NclVacations.ps1 -EmailAlerts -SortByPrice
```

To test without sending email:

```bash
pwsh -NoProfile -File ./Monitor-NclVacations.ps1 -SortByPrice
```

Custom alert values:

```bash
pwsh -NoProfile -File ./Monitor-NclVacations.ps1 \
  -EmailAlerts \
  -AbsoluteAlertThreshold 375 \
  -DropPercentThreshold 8 \
  -RecentLowDays 30 \
  -SortByPrice
```

## Install the user systemd timer

The supplied timer checks at 8 AM, 10 AM, noon, 2 PM, 4 PM, 6 PM, 8 PM and
10 PM. It adds a random delay of up to five minutes so the requests are not
always fired at precisely the same second.

```bash
mkdir -p ~/.config/systemd/user
cp systemd/ncl-price-search.service ~/.config/systemd/user/
cp systemd/ncl-price-search.timer ~/.config/systemd/user/

loginctl enable-linger "$USER"

systemctl --user daemon-reload
systemctl --user enable --now ncl-price-search.timer
```

Check it:

```bash
systemctl --user status ncl-price-search.timer
systemctl --user list-timers | grep ncl
journalctl --user -u ncl-price-search.service -n 100 --no-pager
```

Run immediately:

```bash
systemctl --user start ncl-price-search.service
```

## Data location

By default the Pi monitor writes:

```text
data/NCL-Vacation-Price-History.csv
```

inside the repo. You can point it at an existing CSV with `-CsvPath` if you
want to seed the Pi with your Windows history.

## Future mailing-list version

The monitor deliberately separates deal detection from delivery. Today the
alert event goes to Fastmail. A future JAXPORT service can replace the
single-recipient SMTP step with a bulk-email provider while leaving NCL price
collection and deal detection unchanged.
