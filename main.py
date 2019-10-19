				####################################################################
				####################################################################
				### BEFORE RUNNING CODE, GO TO GMAIL SETTINGS AND LOWER SECURITY ###
				####################################################################
				####################################################################

#importing libraries
import xlrd
import smtplib

loc = './Excel_File.xlsx'  #setting the path of the Excel file
wb=xlrd.open_workbook(loc) #creating object of the workbook
sheet = wb.sheet_by_index(0) #creating object for the sheet at index 0
for i in range (sheet.nrows):
	for i=0						#ignoring first row as it contains heading
	continue					#line (15,16) optional
	group = sheet.cell_value(i,1) #Selecting column 1
	name = sheet.cell_value(i,2)  #Selecting column 2
	Mail-ID= sheet.cell_value(i,3) #Selecting column 3
	remarks= sheet.cell_value(i,4) #Selecting column 4
	try:
		#sending Mail
		d = smtplib.SMTP('smtp.gmail.com', 587)
		d.starttls()
		d.login("Your_Email_ID","Your_Password")    
		message1 = f"Hi {name},\n\n{remarks}\n\nThanks,\n{group}"
		d.sendmail("Your_Email_ID",Mail-ID, message1)
		print(f'Mail Successfully Sent To Mail-ID: {Mail-ID}') #reporting if successful
		d.quit()
	except:
		print(f'Mail Not Sent To Mail-ID: {Mail-ID}') #reporting if Failed

		
					##########################################
					################CODE ENDS ################
					##########################################
