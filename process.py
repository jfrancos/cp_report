import email
from lxml import html 		# pip install lxml
from io import StringIO
import csv
import ldap 				# pip install python_ldap
from dateutil.relativedelta import relativedelta
from dateutil.parser import parse
from datetime import *;
import math
from collections import ChainMap

def td(element, index):
	return element.xpath('td[' + str(index) + ']')[0].text_content().strip()

def format_time(time):  # turns e.g. "10.2 months" into "2018-10-20 / 10.2 months"
	if not time:
		return ""
	time0 = float(time.split()[0])
	time1 = time.split()[1]

	units = [('yr', 'years'), ('hr', 'hours'), ('month', 'months'), ('day', 'days'), ('min', 'minutes')]

	for unit in units:
		if time1.startswith(unit[0]):
			time1 = unit[1]
			break

	ceil_args = {time1: -math.ceil(time0)}
	floor_args = {time1: -math.floor(time0)}

	ceil = (today + relativedelta(**ceil_args)).timestamp()
	floor = (today + relativedelta(**floor_args)).timestamp()

	decimal = time0 % 1
	fraction = (ceil - floor) * decimal
	timestamp = floor + fraction
	newtime = datetime.fromtimestamp(timestamp).strftime("%F")
	return newtime + " / " + time

def ldap_search(uids, attrs):
	ldap_limit = 100
	ldap_db = ldap.initialize("ldaps://ldap.mit.edu:636")
	chunked_uids = [uids[i:i + ldap_limit] for i in range(0, len(uids), ldap_limit)]
	result = []
	for uid_chunk in chunked_uids:
		print(f"querying ldap server for {len(uid_chunk)} kerbs")
		filter = "(|(uid=" + ")(uid=".join(uid_chunk) + "))"
		result += ldap_db.search_s("dc=mit,dc=edu", ldap.SCOPE_SUBTREE, filter, set(attrs + ['uid']))
	result = [item[1] for item in result]
	result = [{key:" / ".join([item.decode() for item in value]) for (key,value) in userdict.items() } for userdict in result ]
	result = [ {item['uid']:item} for item in result ]
	return dict(ChainMap(*result))

## Extract html text from raw multipart email into 'body'
print("Extracting HTML text from raw source")
with open("cpraw") as file:
	msg = email.message_from_file(file)
for part in msg.walk():       
	if part.get_content_type() == "text/html":
		body = part.get_payload(decode=True)
		body = body.decode()
		break

## Convert html text to an lxml tree
print("Contructing lxml tree")
tree = html.parse(StringIO(body), html.HTMLParser())

## Get the relevant parts of the tree
print("Finding relevent nodes")
path = tree.xpath('//table//table//table//table[position() mod 2 = 1]/tbody/tr')

## Date math should be relative to mail header's date
today = parse(msg.get('Date'))
print(f"Email header date is {today}")

kerbs = []
for element in path:
	kerb = td(element, 1) if td(element, 1) else kerb
	kerbs += [kerb]


userinfo = ldap_search(list(set(kerbs)), ['cn', 'roomNumber'])
for kerb in set(kerbs):
	userinfo[kerb]['count'] = kerbs.count(kerb)

kerb_fn = "Username"
archive_fn = "Source"
percent_fn = "Backed Up %"
completed_fn = "Last Completed"
activity_fn = "Last Activity"
cn_fn = "Name"
room_fn = "Room"

print(f"Writing {len(kerbs)} rows for {len(set(kerbs))} distinct users")
with open('cpReport.csv', 'w', newline='') as csvfile:
	fieldnames = [kerb_fn, cn_fn, room_fn, archive_fn, percent_fn, completed_fn, activity_fn]
	writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
	writer.writeheader()

	for element in path:
		kerb = td(element, 1) if td(element, 1) else kerb
		# print(kerb)
		archive = td(element, 2).partition('\n')[0]
		percent = td(element, 5)
		completed = format_time(td(element, 6))
		activity = format_time(td(element, 7))
		cn = userinfo[kerb].get('cn')
		room = userinfo[kerb].get('roomNumber')
		count = userinfo[kerb].get('count')
		count = " (" + str(count) + ")" if count > 1 else ""
		writer.writerow({
			kerb_fn: kerb + count,
			cn_fn: cn,
			room_fn: room,
			archive_fn: archive,
			percent_fn: percent,
			completed_fn: completed,
			activity_fn: activity
		})
