import requests
import json
import arrow
import datetime
import csv
import os
import shutil

import pdb


# Shipstation API creds
api_key = '1faba07e6a5845e7a52111ccda1dae23'
api_secret = 'e5e6bdfc7333440a867d2b34ef133b18'
coffee_sizes = ['10oz', '1lb', 'KILO', '3lb', '5lb']

base_endpoint = 'https://ssapi.shipstation.com/'

def make_api_query(endpoint, params):
	api_endpoint = endpoint + '?' + params
	url = '{}{}'.format(base_endpoint, api_endpoint)
	r = requests.get(url, auth=(api_key, api_secret))
	results = r.json()
	return results	

def get_wholesale_customers():
	endpoint = 'customers'
	params = 'tagId=49965'
	results = make_api_query(endpoint, params)
	return results.get('customers')

def most_recent_week():
	last_week_num = 21
	# d = '%d-W%d' % (2017, datetime.date(2017, 5, 29).isocalendar()[1])
	return last_week_num

def current_week_num():
	return datetime.datetime.now().isocalendar()[1]

def week_num_to_dates(last_week_num):
	# Convert the week number to special string format
	week_num = "%d-W%d" % (2017, last_week_num)

	# Get the start date, and subtract 3 days to make it the previous Thursday at midnight
	start_date = (datetime.datetime.strptime(week_num + '-0', "%Y-W%W-%w") + datetime.timedelta(days=-3)).replace(hour=0, minute=0, second=0, microsecond=0)

	# Get the end date by adding 6 days to the start date, making it Wednesday at 11:59pm
	end_date = (start_date + datetime.timedelta(days=6)).replace(hour=23, minute=59, second=59, microsecond=0)

	return str(start_date), str(end_date)

def orders_by_customer(customer, start_date, end_date):
	# Query for orders by customer for the current week
	endpoint = 'orders'
	params = 'orderDateStart=' + start_date +  '&orderDateEnd=' + end_date + '&customerName=' + customer.get('name')

	order_results = make_api_query(endpoint, params)

	return order_results

def coffee_sizes_in_orders(orders):
	coffee_dict = {}

	def item_to_dict(item):
		for curr_size in coffee_sizes:
			if curr_size in item.get('sku'):
				if not coffee_dict.get(curr_size):
					coffee_dict[curr_size] = 0
				coffee_dict[curr_size] += item.get('quantity')

	for order in orders.get('orders'):
		for item in order.get('items'):
			if 'CFE' in item.get('sku'):
				item_to_dict(item)

	return coffee_dict

def coffee_sizes_to_pounds(coffee_size_dict):
	total_pounds = 0.0
	for size, quantity in coffee_size_dict.iteritems():
		if size == '10oz':
			total_pounds += (quantity*10/16)
		elif size == '1lb':
			total_pounds += quantity
		elif size == 'KILO':
			total_pounds += (quantity*2.2)
		elif size == '3lb':
			total_pounds += (quantity*3)
		elif size == '5lb':
			total_pounds += (quantity*5)
	return total_pounds

def populate_spreadsheet():
	pass

# Query to get all wholesale customers
# Get the most recent "week number" from the spreadsheet
# Convert the week number to start-date / end-date format, to enable querying
# Loop through dates / week numbers until current week is reached
	# Loop through customers
		# Query for all orders for that customer from the start-end date
		# Loop through orders
			# Make a temp dict with coffee bag sizes as keys, quantities as values
		# Convert coffee bag sizes dict to "total pounds"
		# Insert total pounds into spreadsheet

customers = get_wholesale_customers()
most_recent_week_num = most_recent_week()
curr_week_num = current_week_num()

# Loop through all weeks until this week
while (curr_week_num > most_recent_week_num):
	# Get start/end date for week number
	start_date, end_date = week_num_to_dates(most_recent_week_num)

	for customer in customers:
		# Get the shipments for the customer in the specified week
		orders = orders_by_customer(customer, start_date, end_date)
		
		# Get a dict of quantities of all coffee sizes in shipment
		coffee_size_dict = coffee_sizes_in_orders(orders)
		
		# Convert coffee size quantities to a single total pounds number
		customer_pounds = coffee_sizes_to_pounds(coffee_size_dict)

		if customer_pounds > 0:
			print "%s - Week # %d - %.2f" % (customer.get('name'), most_recent_week_num, customer_pounds)

	most_recent_week_num += 1




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