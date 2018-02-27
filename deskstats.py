from __future__ import print_function

import matplotlib.pyplot as plt
import subprocess
import sys
#import pygal
import numpy as np

from counter.counter import Counter
from datetime import date
from email.utils import make_msgid
from python_rt.rt import Rt

# Python 2/3 compatibility
if sys.version_info[0] == 3:
    from email.mime.multipart import MIMEMultipart
    from email.mime.image import MIMEImage
    from email.mime.text import MIMEText
    from tempfile import TemporaryDirectory, TemporaryFile
else:
    import shutil

    from email.MIMEMultipart import MIMEMultipart
    from email.MIMEImage import MIMEImage
    from email.MIMEText import MIMEText
    from tempfile import mkdtemp, TemporaryFile

    class TemporaryDirectory(object):
        """Context manager for tempfile.mkdtemp() to duplicate behavior of Python3, letting me using it with "with" keyword."""
        def __enter__(self):
            self.name = mkdtemp()
            return self.name

        def __exit__(self, exc_type, exc_value, traceback):
            shutil.rmtree(self.name)


# No login credentials needed/given because they're in ~/.netrc.
# For some reason, I can't get login to work any other way. Look into this later, maybe?
rt = Rt(url='https://support.circuit5.org/rt/REST/1.0/',
        skip_login=True)

ticket_open = 'Status != "resolved" AND Status != "rejected"'
stale = 'LastUpdated < "-30 days"'

searches = [
    'Created > "-1 week"',
    'Status = "resolved" AND Resolved > "-1 week"',
    'Updated > "-1 week"',
    'Created > "-1 week" AND %s' % ticket_open,
    ticket_open,
    '%s AND %s' % (ticket_open, stale),
]

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

html = ['<html>',
        '<head>',
        '  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">',
        '</head>',
        '<body bgcolor="#FFFFFF" text="#000000">']

msg = MIMEMultipart()
msg['To'] = 'David Pflug <dpflug@circuit5.org>'
#msg['To'] = 'Court Tech Dept <court_technology_department@circuit5.org>'
msg['From'] = 'David Pflug <dpflug@circuit5.org>'
msg['Subject'] = 'Technology Support Staff Week in Review %s' % date.today()

def make_table(key, data):
    '''Assumes 2-column data because I don't need anything else.'''
    html = ['<table>',
            '  <tr>',
            '    <th>%s</th>' % 'County' if key == 'CF.{County}' else key,
            '    <th>Tickets</th>',
            '  </tr>']
    for i, (c1, c2) in enumerate(data):
        if i % 2 == 0:
            html.append('  <tr style="background:#CCC;">')
        else:
            html.append('  <tr>')
        html.append('    <td style="text-align:right;padding-right:0.5em">%s</td>' % c1)
        html.append('    <td>%s</td>' % c2)
        html.append('  </tr>')

    html.append('</table>')
    html.append('<br>') # We need space below our tables
    html.append('<br>')

    return '\n'.join(html)

img_cids = []
imgs = []

results = []

for search in searches:
    # Do our searches, collecting the results
    results.append(rt.search(Queue = 'Technology Support', order='Owner', raw_query=search))

with TemporaryDirectory() as tempdir:
    # List of images to add after the fact, so we don't have all our graphs appearing at the top of the email
    imgs = []
    # Construct our email
    for sec, search, key, result_list in zip(email_sections, searches, search_keys, results):
        '''Assemble our email sections'''
        ## Count up our ticket totals
        totals = Counter()
        for ticket in result_list:
            tkey = ticket.get(key, 'None')
            if tkey == '':
                tkey = 'None'
            if ',' in tkey:  # Mainly happens when multiple counties are assigned the same ticket
                for subkey in tkey.split(','):
                    totals[subkey] += 1
            else:
                totals[tkey] += 1

        sorted_totals = sorted(totals.items(), key=lambda x: x[1], reverse=("tickets by owner" in sec.lower()))

        ## DEBUG/printed stats
        #print(sorted_totals, file=sys.stderr)

        ## Create bar chart
        xindex = np.arange(len(sorted_totals))
        padding = 1
        chart = plt.bar(xindex + padding,
                        list(x[1] for x in sorted_totals),
                        align='center',)

        # Label our bars
        ax = plt.gca()
        (y_bottom, y_top) = ax.get_ylim()
        y_height = y_top - y_bottom
        for bar in chart:
            height = bar.get_height()
            p_height = (height / y_height)

            if p_height > 0.98:
                label_pos = height - (y_height * 0.05)
            else:
                label_pos = height + (y_height * 0.01)

            ax.text(bar.get_x() + bar.get_width()/2.0,
                    label_pos,
                    '%d' % int(height),
                    ha='center',
                    va='bottom',)

        # Various prettiness
        plt.title("%s\n%s" % (sec, search))
        plt.margins(0.2)
        plt.subplots_adjust(bottom=0.3)
        plt.xticks(xindex + padding,
                   list(x[0] for x in sorted_totals),
                   rotation=90)

        # Render bar charts
        with TemporaryFile(mode='w+b', dir=tempdir) as tf:
            plt.savefig(tf)
            tf.flush()
            tf.seek(0)  # Make sure we're ready for the read() below
            cid = make_msgid()[1:-1]
            html.append('<img src="cid:%s" />' % cid)
            img = MIMEImage(tf.read(), _subtype='png')
            img.add_header('Content-ID', '<%s>' % cid)
            imgs.append(img)

        # Clean up after Matplotlib
        plt.clf()
        plt.cla()

        html.append(make_table(key, sorted_totals))

    html.append('</body>')
    html.append('</html>')

    msg.attach(MIMEText('\n'.join(str(e) for e in html), _subtype='html'))

    for img in imgs:
        msg.attach(img)

    print(msg.as_string())
