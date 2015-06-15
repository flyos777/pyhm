from pprint import pprint
import json, xlwt, random

output_fields = ['performance','competence','talent','workcit','attrition','role_profile','operating_profile','time_in_role']
output_settings = 	{
						'performance':{'name':'Performance Profile','order':1},
						'competence':{'name':'Competence Profile','order':2},
						'talent':{'name':'Talent Profile','order':3},
						'workcit':{'name':'Workman/Citizenship','order':4},
						'attrition':{'name':'Attrition Risk','order':5},
						'role_profile':{'name':'Role Profile','order':8},
						'operating_profile':{'name':'Personal Profile vs Role','order':6},
						'time_in_role':{'name':'Time in role','order':7}
					}


users_dict = {}
users_dict["George Washington"] = {'lname':'Washinghton','fname':'George','company':'wtd','manager':'John Doe','time_in_role':'2-3','competence':'C3','performance':'IN','attrition':'L','talent':'St','workcit':'NGuy','role_profile':'Co','operating_profile':'IM'}
users_dict["Abraham Lincoln"] = {'lname':'Lincoln','fname':'Abraham','company':'dhdh','manager':'John Doe','time_in_role':'2-3','competence':'C3','performance':'O','attrition':'VH','talent':'UC','workcit':'NGuy','role_profile':'Co','operating_profile':'IM'}

managers = ['John Doe', 'Calamity Jane', 'Bernard Shaw']

roles_list = ['Su', 'Ad', 'Co', 'IM', 'OD', 'PD', 'SD']
performance_list = ['O', 'E', 'HS', 'GS', 'IN', 'U', 'NE']
competence_list = ['S1', 'S2', 'S3', 'C1', 'C2', 'C3', 'E1', 'E2', 'E3', 'Sp', 'G', 'NE']
workcit_list = ['RM', 'Emerg', 'InB', 'NGuy', 'PriDo', 'Strug', 'NE']
talent_list = ['St', 'RS', 'HP', 'KC', 'CC', 'OC', 'MC', 'UC', 'NE']
tir_list = ['<1', '1-2', '2-3', '3-4', '>4']
attrition_list = ['L', 'M', 'H', 'VH']

for i in range(1000):
	users_dict["User %s" % i] = {'lname':'Doe','fname':'John','company':'dhdh','manager':'%s' % random.choice(managers),'time_in_role':'%s' % random.choice(tir_list),'competence':'%s' % random.choice(competence_list),'performance':'%s' % random.choice(performance_list),'attrition':'%s' % random.choice(attrition_list),'talent':'%s' % random.choice(talent_list),'workcit':'%s' % random.choice(workcit_list),'role_profile':'%s' % random.choice(roles_list),'operating_profile':'%s' % random.choice(roles_list)}

# pprint('')

# Write dictionnary to json file
# with open('settings.json', 'w') as f:
# 	json.dump(data, f, separators=(',', ':'),indent=3)

json_data=open('settings.json').read()
data = json.loads(json_data)


def get_user_dict(user):
	# Set Output dictionnary
	output = {}
	user_infos = {}
	
	#Set informations to use in profile
	time_in_role = users_dict[user]['time_in_role']
	
	try:
		total = 0
		for field in output_fields: # Loop for fields into user
			
			try: # Try to set rating* variables for fields

				rating = users_dict[user][field]
				
				### Set Custom settings for fields variables ###
				if field == "competence":
					rating_color = data['%s_rating' % field][rating][time_in_role]

				elif field == "role_profile" or field == "time_in_role":
					rating_color = "White"

				elif field == "operating_profile":
					rating = users_dict[user]['operating_profile']
					dif_role_op = int(data['role_profile_rating'][users_dict[user]['role_profile']]) - int(data['role_profile_rating'][users_dict[user]['operating_profile']])
					rating_color = data['role_profile_rating']['scores'][str(dif_role_op)]
					if rating_color == "Green":
						rating_score = 0
					else:
						rating_score = 2
				
				else: # Else, not concerned by custom settings, set default
					rating_color = data['%s_rating' % field][str(rating)]
					rating_score = data['%s_rating' % field]['scores'][rating_color]

				### Set custom fucked up setting ###
				if field == "competence":
					rating_score = (1.25 * rating_score)

				### Write key,values into user directory for each field
				user_infos['%s_rating' % field] = rating
				user_infos['%s_rating_color'% field] = rating_color
				user_infos['%s_rating_score'% field] = rating_score
				
			except: # If can't set fields, set default one, to keep the machine working...
				rating_score = 0
				user_infos['%s_rating_score'% field] = rating_score
			
			total += rating_score #Increment total score for each field

		if total < 3: # Set total color
			total_color = "Green"
		elif total <= 4:
			total_color = "Amber"
		elif total > 4:
			total_color = "Red"
		
		### Write default user informations ###
		user_infos['company'] = users_dict[user]['company']
		user_infos['manager'] = users_dict[user]['manager']
		user_infos['global_profiling'] = total_color
		user_infos['global_profiling_score'] = total
		
		output[user] = user_infos # Set dictionnary with name as main key
		return output
	
	except: #If no user found, pass
		pass

def get_all_users_dict(**kwargs):

	output = []

	for user in users_dict: # Loop for users
		sorted(users_dict.items(), key = lambda tup: (tup[1]["fname"]))
		if len(kwargs) > 0: #If arguments
			count = 0
			for name, value in kwargs.items():
				if get_user_dict(user)[user][name] == value:
					count += 1
			if len(kwargs) == count:
				output.append(get_user_dict(user))

		else:
			output.append(get_user_dict(user))
	
	return output

def create_xls():
	# Init excel file and sheet
	workbook = xlwt.Workbook(style_compression=2)
	sheet = workbook.add_sheet("Sheet Name")

	#sheet.col(0).width = 256 * (len(key) + 1)

	def set_style(color):
		headers = []
		# https://docs.google.com/spreadsheets/d/1ihNaZcUh7961yU7db1-Db0lbws4NT24B7koY8v8GHNQ/pubhtml?gid=1072579560&single=true
		font_color = "white"
		bold = 0
		if color == "Amber":
			color = "light_orange"
		elif color == "Red":
			color = "red"
		elif color == "Green":
			color = "lime"
		elif color == "White":
			color = "white"
			font_color = "black"
		elif color == "Bold":
			color = "black"
			bold = 1

		style = xlwt.easyxf('font: color %s, bold %s; pattern: pattern solid, fore_color %s; borders: left_colour white, right_colour white, top_colour white, bottom_colour white, top thin, right thin, bottom thin, left thin;' % (font_color,bold,color))
		return style


	####### HEADER  PLACE ######
	row = 5 # Starting row for generated fields

	sheet.write(row-1, 0, "TM Name",set_style("Bold")) #Write Default rows
	sheet.write(row-1, 1, "Manager",set_style("Bold"))
	sheet.write(row-1, 2, "Global Profiling",set_style("Bold"))
	for title in output_fields:
		sheet.write(4, output_settings[title]['order']+2, "%s %s" % (output_settings[title]['order']+2,output_settings[title]['name']),set_style("Bold"))
	


	####### USERS INFOS PLACE ######
	users_d = get_all_users_dict() #Write for fields
	if len(users_d ) > 0:
		for user_dict in users_d : # For each user in dictionnary

			for user in user_dict: # For each dictionnary in the user's dictionnary (should have only one) // Please change it!
				
				#sorted(users_dict.items() , key = lambda tup: (tup[1]["global_profiling_score"]))
				### Default infos, name,manager,globale_profiling ###
				sheet.write(row, 0, user,set_style("White") )
				sheet.write(row, 1, user_dict[user]['manager'],set_style("White"))
				sheet.write(row, 2, "",set_style(user_dict[user]['global_profiling'])) # Global Profiling
				
				### Dynamic fields ###
				for field in output_fields: 
					sheet.write(row, output_settings[field]['order']+2, "%s" % (user_dict[user]['%s_rating' % field]),set_style(user_dict[user]['%s_rating_color' % field]))

			row += 1

		workbook.save("foobar.xls")
	else:
		print "Not entry found!"

if  __name__ == "__main__": # Windows only
	pass
	#pprint(get_user_dict("George Washington"))
	create_xls()
