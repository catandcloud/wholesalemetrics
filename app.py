import requests
import json
import arrow
import csv
import os
import shutil



# Shipstation API creds
api_key = '1faba07e6a5845e7a52111ccda1dae23'
api_secret = 'e5e6bdfc7333440a867d2b34ef133b18'

base_endpoint = 'https://ssapi.shipstation.com/'

# End the order query at midnight on the current day
start_date = str(arrow.utcnow().to('US/Pacific'))
end_date = str(arrow.utcnow().to('US/Pacific'))
# end_date = str(arrow.utcnow().to('US/Pacific').replace(hour=0, minute=0, second=0, microsecond=0))

endpoint = 'orders/listbytag'
params = 'shipDate>=' + start_date + '&shipDate<=' + end_date + '&orderStatus=shipped' + '&tagId=49965'
api_endpoint = endpoint + '?' + params

url = '{}{}'.format(base_endpoint, api_endpoint)
r = requests.get(url, auth=(api_key, api_secret))

order_results = r.json()



coffees = {}

# Loop through orders and get the coffees
for order in order_results['orders']:
	# Skip orders that were canceled or already shipped
	if order.get('orderStatus') in ['shipped', 'cancelled']:
		continue

	# Loop through items and add to coffees dict according to sku
	for item in order['items']:
		
		# Get sku and quantity of current item
		sku = item.get('sku')
		quantity = item.get('quantity')
		
		# Break if there is an item that does not have a sku - it wouldn't be possible to account for in the spreadsheet
		if not sku:
			print "No SKU for item " + item.get('name')
			continue

		# We only want coffee skus
		if "CFE" not in sku:
			continue

		# We only want the following coffee skus, so skip all others
		if not any(x in sku for x in ['10oz', '1lb', 'KILO', '5lb']):
			continue

		# Split the sku at the grind (since it has nothing to do with this)
		sku = sku.split('-')[1] + '-' + sku.split('-')[2]

		# Treat all subscription skus as the same
		if 'Sub' in sku or 'Staff' in sku:
			sku = 'Sub' + '-' + sku.split('-')[1]

		# Init the sku coffees dict entry for each new sku
		if not coffees.get(sku):
			coffees[sku] = 0
		
		# Increment the coffee sku entry by the quantity in the current order
		coffees[sku] += quantity