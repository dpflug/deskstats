#!/usr/bin/env python
import argparse
import pygal
import subprocess
import sys
import tempfile

from collections import Counter
from datetime import date
from email.message import EmailMessage
from email.utils import make_msgid
from python_rt.rt import Rt

parser = argparse.ArgumentParser(description='Generate CTSS "Week in Review" email')
parser.add_argument('--date', '-d',
                    help='Select the week to generate stats for. (YYYY-MM-DD required) (Also, not yet working)')

args = parser.parse_args()

rt = Rt('https://support.circuit5.org/rt/REST/1.0/', 'dpflug', subprocess.check_output(['pass', 'Work'], timeout=30).decode('utf-8').split('\n')[0], skip_login=True)

ticket_open = 'Status != "resolved" AND Status != "rejected"'
stale = 'LastUpdated < "-30 days"'

searches = [
    'Created > "-1 week"',
    'Status = "resolved" AND Resolved > "-1 week"',
    'Updated > "-1 week"',
    f'Created > "-1 week" AND {ticket_open}',
    ticket_open,
    f'{ticket_open} AND {stale}']

search_keys = [
    'CF.{County}',
    'Owner',
    'Owner',
    'Owner',
    'Owner',
    'Owner']

email_sections = [
    'Tickets opened last week by County',
    'Tickets resolved last week by Owner',
    'Tickets worked on last week by Owner',
    'Tickets opened and not resolved last week by Owner',
    'Open tickets by Owner',
    'Stale tickets by Owner']

today = date.today()

html = ['<html>',
        '<head>',
        '  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">',
        # For SVG/PNG alternates
        #'  <style>',
        #'    .show {',
        #'      height: 100% !important;',
        #'      width: 100% !important;',
        #'    }',
        #'',
        #'    .hide {',
        #'      display: none;',
        #'    }',
        '  </style>',
        '</head>',
        '<body bgcolor="#FFFFFF" text="#000000">']

msg = EmailMessage()
msg['Subject'] = f'Technology Support Staff Week in Review {today}'
msg['From'] = 'David Pflug <dpflug@circuit5.org>'
msg['To'] = 'Court Tech Dept <court_technology_department@circuit5.org>'
#msg['To'] = 'David Pflug <dpflug@circuit5.org>'
msg.preamble = "This is a multi-part message in MIME format. Please enable html email in your client to view it."
msg.set_content("This is a multi-part message in MIME format. Please enable html email in your client to view it.")

def make_table(key, data):
    '''Assumes 2-column data because I don't need anything else.'''
    html = ['<table>',
            '  <tr>',
            f'    <th>{key}</th>',
            '    <th>Tickets</th>',
            '  </tr>']
    for i, (c1, c2) in enumerate(data):
        if i % 2 == 0:
            html.append('  <tr style="background:#CCC;">')
        else:
            html.append('  <tr>')
        html.append(f'    <td style="text-align:right;padding-right:0.5em">{c1}</td>')
        html.append(f'    <td>{c2}</td>')
        html.append('  </tr>')

    html.append('</table>')
    html.append('<br>') # We need space below our tables
    html.append('<br>')
    
    return '\n'.join(html)

img_cids = []
imgs = []

results = []

for search in searches:
    results.append(rt.search(Queue = 'Technology Support', order='Owner', raw_query=search))

with tempfile.TemporaryDirectory() as tempdir:
    pngs = []
    #svgs = []
    for sec, search, key, result_list in zip(email_sections, searches, search_keys, results):
        '''Assemble our email sections'''
        # Count up our ticket totals
        totals = Counter()
        for ticket in result_list:
            tkey = ticket.get(key, 'None')
            if tkey == '':
                tkey = 'None'
            totals[tkey] += 1

        sorted_totals = sorted(totals.items(), key=lambda x: x[1], reverse=("tickets by owner" in sec.lower()))
        
        #DEBUG
        print(sorted_totals, file=sys.stderr)

        # Create bar charts
        chart = pygal.Bar(x_label_rotation=30,
                          show_legend=False,
                          rounded_bars=10,
                          print_values=True,
                          print_zeroes=False)
        chart.title = f'{sec}\n{search}'
        chart.x_labels = [x[0] for x in sorted_totals]
        chart.add('Tickets', [x[1] for x in sorted_totals])

        # Render bar charts
        _, tf = tempfile.mkstemp(dir=tempdir)
        chart.render_to_png(tf)
        cid = make_msgid()[1:-1]
        #html.append(f'<img class="hide" src="cid:{cid}" />')
        html.append(f'<img src="cid:{cid}" />')
        pngs.append((tf, cid))

        # SVG in email is poorly supported
        #h, tf = tempfile.mkstemp(dir=tempdir)
        #print(dir(h))
        #chart.render_to_file(tf)
        #cid = make_msgid()[1:-1]
        #html.append(f'<img class="show" width="0" height="0" src="cid:{cid}" />')
        #svgs.append((tf, cid))

        html.append(make_table(key, sorted_totals))

    html.append('</body>')
    html.append('</html>')
        
    msg.add_alternative('\n'.join(str(e) for e in html), subtype='html')

    #for png, svg in zip(pngs, svgs):
    for png in pngs:
        with open(png[0], 'rb') as i:
            msg.get_payload()[1].add_related(i.read(),
                                             'image',
                                             'png',
                                             cid=png[1])
        #with open(svg[0], 'rb') as i:
        #    msg.get_payload()[1].add_related(i.read(),
        #                                     'image',
        #                                     'svg+xml',
        #                                     cid=svg[1])

    print(msg)
