#!/usr/bin/env python
import argparse
import re
import requests
import subprocess

from datetime import date,timedelta
from email.message import EmailMessage
from email.headerregistry import Address
from email.utils import make_msgid

parser = argparse.ArgumentParser(description='Generate CTSS "Week in Review" email')
parser.add_argument('--date', '-d',
                    help='Select the week to generate stats for. (YYYY-MM-DD required)')

args = parser.parse_args()

rt_url = 'https://support.circuit5.org/rt/'
rt_search = rt_url + 'Search/Chart.html'

tech_queue = 'Queue = "Technology Support"'
ticket_open = 'Status != "resolved" AND Status != "rejected"'
stale = 'LastUpdated < "-30 days"'
common = {'SavedChartSearchId': 'new', 'ChartStyle': 'bar'}
by_county = {'PrimaryGroupBy': 'CF.{1}'}
by_user = {'PrimaryGroupBy': 'Owner.Name'}

if not args.date:
    print("Using today.")
    searches = [
        # Opened last week
        {**{'Query': '{} AND Created > "-1 week"'.format(tech_queue)},
         **common,
         **by_county},
        # Resolved last week
        {**{'Query': '{} AND Status = "resolved" AND Resolved > "-1 week"'.format(tech_queue)},
         **common,
         **by_user},
        # Worked on last week
        {**{'Query': '{} AND Updated > "-1 week"'.format(tech_queue)},
         **common,
         **by_user},
        # Opened and not resolved
        {**{'Query': '{} AND Created > "-1 week" AND {}'.format(tech_queue, ticket_open)},
         **common,
         **by_user},
        # Open by owner
        {**{'Query': '{} AND {}'.format(tech_queue, ticket_open)},
         **common,
         **by_user},
        # Not updated in over 30 days
        {**{'Query': '{} AND {} AND {}'.format(tech_queue, ticket_open, stale)},
         **common,
         **by_user},
    ]
else:
    Y, M, D = args.date.split('-')
    sunday = date(int(Y), int(M), int(D))
    if sunday.weekday() != 6:
        sunday = sunday - timedelta(days=sunday.weekday() + 1)

    saturday = sunday + timedelta(days=6)
    print(sunday, saturday)
    searches = [
        # Opened last week
        {**{'Query': '{} AND Created >= "{}" AND Created <= "{}"'.format(tech_queue, sunday, saturday)},
         **common,
         **by_county},
        # Resolved last week
        {**{'Query': '{} AND Status = "resolved" AND Resolved >= "{}" AND Resolved <= "{}"'.format(tech_queue, sunday, saturday)},
         **common,
         **by_user},
        # Worked on last week
        {**{'Query': '{} AND Updated >= "{}" AND Updated </ "{}"'.format(tech_queue, sunday, saturday)},
         **common,
         **by_user},
        # Opened and not resolved
        {**{'Query': '{} AND Created >= "{}" AND Created <= "{}" AND {}'.format(tech_queue, sunday, saturday, ticket_open)},
         **common,
         **by_user},
        # Open by owner
        {**{'Query': '{} AND {}'.format(tech_queue, ticket_open)},
         **common,
         **by_user},
        # Not updated in over 30 days
        {**{'Query': '{} AND {} AND {}'.format(tech_queue, ticket_open, stale)},
         **common,
         **by_user},
    ]



msg = EmailMessage()
msg['Subject'] = "Week in Review"
msg['From'] = Address('David Pflug', 'dpflug', 'circuit5.org')
msg['To'] = Address('David Pflug', 'dpflug', 'circuit5.org')
msg.set_content("This email is html-only. Please enable html email in your client.")

email_headers = [
    "<p><b>Tickets opened last week by County</b></p>\n<div>",
    "</div>\n<p><b>Tickets resolved last week by Owner</b></p>\n<div>",
    "</div>\n<p><b>Tickets worked on last week by Owner</b></p>\n<div>",
    "</div>\n<p><b>Tickets opened and not resolved last week by Owner</b></p>\n<div>",
    "</div>\n<p><b>Open tickets by Owner</b></p>\n<div>",
    "</div>\n<p><b>Stale tickets by Owner</b></p>"
]

s = requests.Session()
headers = { 'user-agent': 'week_in_review.py',
            'referer': '/rt/' }

email = ['<html>',
         '<head>',
         '<meta http-equiv="Content-Type" content="text/html; charset=utf-8">',
         '</head>',
         '<body bgcolor="#FFFFFF" text="#000000">']

img_cids = []
imgs = []

url_search = re.compile('/rt/Search/Chart.html\?CSRF_Token=.*?(?=")')
img_search = re.compile('(?<=img src=").*(?=")')


auth = ('dpflug',
        subprocess.check_output(['pass', 'Work'], timeout=30)
        .decode('utf-8').split('\n')[0])

def rt_get(search):
    print(search)
    r = s.get(rt_search, params=search, auth=auth, verify=False)
    url2 = url_search.search(r.content.decode('utf-8'))
    if url2:
        print(url2.group(0))
        return s.get(url2.group(0),
                     auth=auth,
                     verify=False)
    else:
        return r
    

def parse_chart_page(search):
    r = rt_get(search)
    print(r.headers)
    headers['referer'] = url

    ts = ""
    tsb = False

    for line in r.iter_lines():
        l = line.decode('utf-8')
        if l.find('img ') >= 0:
            img = l[:l.find('/>')+2]
        if l.find('table class="collection-as-table') >= 0:
            tsb = True
        if tsb:
            ts = ts + l + '\n'
            if l.find('</table>') >= 0:
                break

    #iurl = img_search.search(img).group(0)
    #i = rt_get(iurl).content
    print(img, ts)
    i = None

    return (i, ts)


# There's some session dark magic happening.
# If I don't make 1 authed request first, I get bad data for the first search.
s.get(rt_url, auth=auth, verify=False)

for i, search in enumerate(searches):
    email.append(email_headers[i])
    img, ts = parse_chart_page(search)
    #imgs.append(img)
    #img_cids.append(make_msgid())
    # Note: cids are wrapped in <>
    #email.append('<img src="%s" />' % img_cids[i][1:-1])
    email.append(ts)

email.append('</body>')
email.append('</html>')

msg.add_alternative('\n'.join(email), subtype='html')

for i, img in enumerate(imgs):
    msg.get_payload()[1].add_related(img, 'image', 'png', cid=img_cids[i])

print(msg)
