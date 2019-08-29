import email
from lxml import html
from io import StringIO
import csv
import dateparser
import ldap

## Extract html text from raw multipart email into 'body'
with open("cpraw") as file:
	msg = email.message_from_file(file)
for part in msg.walk():       
	if part.get_content_type() == "text/html":
		body = part.get_payload(decode=True)
		body = body.decode()
		break


## Convert html text to an lxml tree
tree = html.parse(StringIO(body), html.HTMLParser())


## Get the relevant parts of the tree
path = tree.xpath('//table//table//table//table[position() mod 2 = 1]/tbody/tr')


ldap_db = ldap.initialize("ldaps://ldap.mit.edu:636")


def td(element, index):
	return element.xpath('td[' + str(index) + ']')[0].text_content().strip()

def format_time(time):
	if not time:
		return ""
	time0 = str(round(float(time.split()[0])))
	newtime = time0 + " " + time.split()[1].replace("hrs", "hours").replace("yrs", "years") + " ago"
	newtime = dateparser.parse(newtime).strftime("%x")
	newtime = newtime + " / " + time
	return newtime


kerb = ""

kerb_fn = "Username"
archive_fn = "Source"
percent_fn = "Backed Up %"
completed_fn = "Last Completed"
activity_fn = "Last Activity"
cn_fn = "Name"
room_fn = "Room"

with open('cpReport.csv', 'w', newline='') as csvfile:
	fieldnames = [kerb_fn, cn_fn, room_fn, archive_fn, percent_fn, completed_fn, activity_fn]
	writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
	writer.writeheader()

	for element in path:
		kerb = td(element, 1) if td(element, 1) else kerb
		print(kerb)
		result = ldap_db.search_s("dc=mit,dc=edu", ldap.SCOPE_SUBTREE, "uid=" + kerb, ['cn', 'roomNumber'])
		archive = td(element, 2).partition('\n')[0]
		percent = td(element, 5)
		completed = format_time(td(element, 6))

		activity = format_time(td(element, 7))

		cn = result[0][1].get('cn')[0].decode() if result[0][1].get('cn') else ""
		room = result[0][1].get('roomNumber')[0].decode() if result[0][1].get('roomNumber') else ""

		writer.writerow({kerb_fn: kerb, cn_fn: cn, room_fn: room, archive_fn: archive, percent_fn: percent, completed_fn: completed, activity_fn: activity})


