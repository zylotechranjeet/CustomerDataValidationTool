import xlrd
import re
workbook = xlrd.open_workbook('/home/rkverma11254/Downloads/Syndicate_DB_082119.xlsx')
worksheet = workbook.sheet_by_name('Sheet1')
inum = worksheet.nrows+1
print('Total Rows: ',worksheet.nrows)

# Test Cases for Postal Code Verification:
for row in range(worksheet.nrows):
    Postal = str(worksheet.cell_value(row, 11))
    if re.findall("^([0-9A-Za-z..\(\)\s])+([\s\s\-]?)[-0-9A-Za-z0-9..]+([\s\s\-]?)[0-9A-Za-z0-9..]*$",Postal):
        pass
    else:
        with open("/home/rkverma11254/Music/PostalCode.txt", 'a+') as post:
            post.writelines("Row Number: {0} Postal: {1}".format(row, Postal) + "\n")
        post.close()
#Test Cases for Prefix:
# for row in range(worksheet.nrows):
#     prefix = str(worksheet.cell_value(row,18))
#     if re.findall("^(Mr|Ms|MR|MS|MR.|MS.|Mr.|Ms.)$",prefix):
#         pass
#     else:
#         with open("/home/rkverma11254/Music/PrefixInName.txt", 'a') as prefix:
#             prefix.writelines("Row Number: {0} Prefix: {1}".format(row, prefix) + '\n')
#         print("Row Number: {0} Prefix: {1}".format(row, prefix))
#         prefix.close()
#Test Cases for First Name:
# for row in range(worksheet.nrows):
#     First_Name = worksheet.cell_value(row,19)
#     if re.findall("^([A-Za-z'0-9]?)+([\s]?)([\sA-Za-z'0-9])+([\s\s]?)([a-zA-Z'0-9_]*)$",First_Name):
#         pass
#     else:
#         with open("/home/rkverma11254/Music/FirstName.txt", 'a+') as FN:
#             FN.writelines("Row Number: {0} FirstName: {1}".format(row, First_Name) + "\n")
#         FN.close()
# Test Cases for Last_Name:
# for row in range(worksheet.nrows):
#     Last_Name = worksheet.cell_value(row,20)
#     if re.findall("^([A-Za-z'0-9’..ß]?)+([\s]?)([\sA-Za-z'0-9’..ß])+([\s\s]?)([a-zA-Z'0-9’_..ß]*)$",Last_Name):
#         pass
#     else:
#         with open("/home/rkverma11254/Music/LastName.txt", 'a+') as LN:
#             LN.writelines("Row Number: {0} LastName: {1}".format(row, Last_Name) + "\n")
#         LN.close()
#Test Cases for Job Title:
# for row in range(worksheet.nrows):
#     job_title = worksheet.cell_value(row,22)
#     if re.findall("^([a-zA-Z])+([\s\/]?[A-Za-z-&]*)+([\s]?)([A-Za-z]*)$",job_title):
#         pass
#     else:
#         with open("/home/rkverma11254/Music/JobTitle.txt",'a+') as JT:
#             JT.write("Row number: {0} and Job Title: {1}: ".format(row, job_title) + "\n")
#         JT.close()
#Test Cases for Address:
# for row in range(worksheet.nrows):
#     Add1 = worksheet.cell_value(row,9)
#     if re.findall("^([a-zA-Z0-9.,'ı،\s])+([\'\:\.\-\_\,\s\&\"]?)([a-zA-Z0-9.,'\/\"\"]?)+(([\(\'\:\.\-\_\,\s\&\")]?)([a-zA-Z0-9.,']?))+([\'\:\.\-\_\,\s\&\/\#]?)([A-Za-z\"])+([\'\:\.\-\_\,\s\&\/\#]?)([-A-Za-z.,\"])+([\'\:\.\-\_\,\s\&\/\#]?)([A-Za-z0-9\"])+(([-&\/,.#\sA-Za-z0-9\"].?)*)$",Add1):
#         pass
#     else:
#         with open("/home/rkverma11254/Music/Address1.txt", 'a+') as add:
#             add.writelines("Row Number: {0} Address1: {1}".format(row, Add1) + "\n")
#         add.close()
# Test Cases for Website verification:
for row in range(worksheet.nrows):
    website = worksheet.cell_value(row,10)
    if re.findall("^http[s]?://?[a-zA-Z0-9.]*$|^[www].([a-zA-Z0-9.-])*$",website):
        pass
    else:
        with open("/home/rkverma11254/Music/Website.txt", 'a+') as ws:
            ws.writelines("Row Number: {0} WebsiteName: {1}".format(row, website) + "\n")
        ws.close()
#For Getting Invalid Email id:
# for row in range(worksheet.nrows):
#     email = (worksheet.cell_value(row,24))
#     if re.findall('(^[a-zA-Z0-9]+[\.\_\-]?)+([a-zA-Z0-9]+[\.\-\_]?)+(\@[a-zA-Z0-9])+([\.\-\_]?)([a-zA-Z0-9])+([\.\-\_]?)[a-zA-Z0-9.]+([\.\-\_]?)[a-zA-Z0-9.]*$',email):
#         pass
#     else:
#         with open("/home/rkverma11254/Music/Email.txt",'a+') as f1:
#             f1.writelines("Row Number: {0} Email: {1}".format(row,email)+"\n")
#         #print("Row Number: {0} Email address: {1}".format(row,email))
#         f1.close()
# # Test Case for Phone no verification:
# for row in range(worksheet.nrows):
#     phoneNo = (worksheet.cell_value(row,16))
#     regex = "^\+?\d*\s?\(?\d{2}\)?[-.\s]?\d{2}[-.\s]?\d*$"
#     if re.findall(regex,phoneNo):
#         pass
#     else:
#         with open("/home/rkverma11254/Music/PhoneNumber.txt", 'a+') as f1:
#             f1.writelines("Row Number: {0} Phone number: {1}".format(row, phoneNo) + "\n")
#             #print("Row Number: {0} Company name: {1}".format(row, phoneNo))
#         f1.close()
# Test Case for state verification:
for row in range(worksheet.nrows):
    state = str(worksheet.cell_value(row,12))
    regex ="(^[a-zA-Zı',..’-]+[\s]?)+([a-zA-Z0-9..'’]+[\s]?)+([a-zA-Z’..]*)$"
    if re.findall(regex,state):
        pass
    else:
        with open("/home/rkverma11254/Music/StateName.txt", 'a+') as f1:
            f1.writelines("Row Number: {0} State: {1}".format(row, state) + "\n")
            #print("Row Number: {0} State: {1}".format(row, state))
        f1.close()
# # Test Case for Country name verification:
for row in range(worksheet.nrows):
    country = str(worksheet.cell_value(row,9))
    regex = "(^[a-zA-Z,']+[\s]?)+([a-zA-Z0-9']+[\s]?)+([a-zA-Z]*)$"
    list =['United States','United Kingdom','Bulgaria','Costa Rica','Brazil','Switzerland','Denmark','Luxembourg','Sweden','Chile','Netherlands','Canada','Costa Rica','Malawi','Hong Kong','Pakistan','China','Belgium','Mexico','Romania','Czech Republic','New Zealand','Sri Lanka','Belgium','Colombia','Bolivia','Colombia','Korea, Republic Of','Russian Federation'
,'Japan','Indonesia','Uganda','Tanzania, United Republic Of','Cayman Islands','Poland','Finland','France','Germany','Botswana'
,'Hungary','South Africa','Spain','Ukraine','Vietnam','Australia','Czech Republic','India','Italy','Macedonia','Mexico',
'New Zealand','Puerto Rico','Portugal','Romania','Thailand','Trinidad And Tobago','Tanzania, United Republic Of','Bosnia And Herzegovina','Cambodia','Taiwan, Republic Of China','United Arab Emirates','Peru','Turkey','Lebanon','Norway','Malaysia','Malta','Iraq','Albania','Singapore','Jamaica','Philippines','Dominican Republic','Panama','Austria','Egypt','Argentina','Ivory Coast','Israel',"Lao People's Democratic Republic",'Kenya','Qatar','Bhutan','Malta','Bahrain','Cyprus','Mauritius','Saudi Arabia','Lithuania','Senegal','Angola','Azerbaijan','Croatia','Estonia','Georgia','Ireland','Israel','Kazakhstan','Kuwait','Latvia','Libyan Arab Jamahiriya','Moldova','Nigeria','Oman','Slovakia','Togo','Brunei Darussalam','Guatemala','Myanmar','Aruba','Bahamas','Gibraltar','Lesotho','Serbia And Montenegro','Ecuador','Greece','Morocco','Central African Republic','Bermuda','Bangladesh','Afghanistan','Iceland','Serbia','Niger','Macau','Barbados','Congo','Algeria','Madagascar','Slovenia','Belarus'
,'Cameroon','Ghana','Jordan','Rwanda','Uzbekistan','Kyrgyzstan','Venezuela','Tunisia','Papua New Guinea','Honduras','Ethiopia','Belize','French Polynesia','Monaco','WARSZAWA','HUDSON','HARRIS','DANE','NEW YORK','LOS ANGELES','MULTNOMAH','TRAVIS','MARICOPA','ELKHART','JACKSON','CUMBERLAND','SOMERSET','PORTSMOUTH CITY','JEFFERSON','JACKSON','DISTRICT OF COLUMBIA','OKLAHOMA','STEARNS','SUMNER','ERIE','ONTARIO','DAVIDSON'
           ,'COLLIN','SANTA CLARA','FAIRFAX','WAUPACA','FULTON','JOHNSON','CHESTER','MIDDLESEX','SAN FRANCISCO','RICHMOND','DALLAS','HARRIS','SAN DIEGO','WARSZAWA','MONTGOMERY','COOK']
    if re.findall(regex,country) and country in list:
        pass
    else:
        with open("/home/rkverma11254/Music/CountryName.txt", 'a+') as f1:
            f1.writelines("Row Number: {0} Country: {1}".format(row, country) + '\n')
        #print("Row Number: {0} Country: {1}".format(row, country))
        f1.close()
#Test Case for company name verification
for row in range(worksheet.nrows):
    company = str(worksheet.cell_value(row,7))
    if re.findall("^([a-zA-Z0-9.,'áçéãı +.,-])+([\'\:\.\-\_\,\s\&\+\!\–#\/]?)([a-zA-Z0-9.,'ıáçéã.,-]?)+(([\(\'\:\.\-\_\,\s\&\+\–#)]?)([a-zA-Z0-9.,'áçéãı.,-]?))+([\'\:\.\-\_\,\s\&\+\/]?)([a-zA-Z0-9.,'áçéãı.,-]*)",company):
       pass
    else:
        with open("/home/rkverma11254/Music/CompanyName.txt",'a+') as f1:
            f1.writelines("Row Number: {0} Company name: {1}".format(row,company)+"\n")
        f1.close()
# Test Cases for City verification:
for row in range(worksheet.nrows):
    city = str(worksheet.cell_value(row,15))
    if re.findall("(^[a-zA-Zı.'-]+[\s]?)+([a-zA-Z0-9',./-]+[\s]?)+([a-zA-Z]*)|^[a-zA-Zı']+$",city):
        pass
    else:
        with open("/home/rkverma11254/Music/CityName.txt", 'a+') as f1:
            f1.writelines("Row Number: {0} City: {1}".format(row, city) + "\n")
        #print("Row Number: {0} City: {1}".format(row, city))
        f1.close()