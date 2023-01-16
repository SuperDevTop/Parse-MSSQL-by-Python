
import pyodbc
import xml.etree.ElementTree as ET

# for x in pyodbc.drivers():
#     print(x)

conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=DESKTOP-BLIBIF4;'
                      'Database=EDOTXML;'
                      'Trusted_Connection=yes;')

cursor = conn.cursor()
cursor.execute('SELECT * FROM srsdorisv2prod.userreference')

result = cursor.fetchall()

for i in result:
    id = i[0]
    user_id = i[1]
    UserWokFlow_id = i[2]
    ReferenceNo = i[3]
    ReferenceQA = i[4]
    
    flag = True

    try:
        myroot = ET.fromstring(ReferenceQA)
    except:
        flag = False

    data_to_insert = []
    
    if flag is True:
        for x in myroot.findall('formrow'):
            label = x[0].attrib['label']

            if label.find("'") != -1:
                label = label.replace("'", "''")

            data_to_insert.append(label)

            if(x[0].text == None):
                text = "None"
            else:
                text = x[0].text

                if text.find("'") != -1:
                    text = text.replace("'", "''")

            data_to_insert.append(text)
        
        if len(data_to_insert) < 14:
            length = len(data_to_insert)
            for k in range(14 - length):
                data_to_insert.append("None")
    else:
        for j in range(14):
            data_to_insert.append("None")

    query = "insert into srsdorisv2prod.userreferenceParsed(user_id," \
            "UserWorkFlow_id, ReferenceNo, Rcall1, Rcall1Answer, Rcall2," \
            "Rcall2Answer, Rq1, Rq1Answer, Rq2, Rq2Answer, Rq3, Rq3Answer," \
            "Rq4, Rq4Answer, Rq5, Rq5Answer) values(" + str(user_id) + "," + str(UserWokFlow_id) \
            + "," + str(ReferenceNo) + ",'" + data_to_insert[0] + "','" \
            + data_to_insert[1] + "','" + data_to_insert[2] + "','" + data_to_insert[3] + "','" \
            + data_to_insert[4] + "','" + data_to_insert[5] + "','" + data_to_insert[6] + "','" \
            + data_to_insert[7] + "','" + data_to_insert[8] + "','" + data_to_insert[9] + "','" \
            + data_to_insert[10] + "','" + data_to_insert[11] + "','" + data_to_insert[12] + "','" \
            + data_to_insert[13] + "')"

    try:
        cursor.execute(query)
    except:
        data_to_insert = []
        for k in range(14):
            data_to_insert.append("None")
        
        query = "insert into srsdorisv2prod.userreferenceParsed(user_id," \
            "UserWorkFlow_id, ReferenceNo, Rcall1, Rcall1Answer, Rcall2," \
            "Rcall2Answer, Rq1, Rq1Answer, Rq2, Rq2Answer, Rq3, Rq3Answer," \
            "Rq4, Rq4Answer, Rq5, Rq5Answer) values(" + str(user_id) + "," + str(UserWokFlow_id) \
            + "," + str(ReferenceNo) + ",'" + data_to_insert[0] + "','" \
            + data_to_insert[1] + "','" + data_to_insert[2] + "','" + data_to_insert[3] + "','" \
            + data_to_insert[4] + "','" + data_to_insert[5] + "','" + data_to_insert[6] + "','" \
            + data_to_insert[7] + "','" + data_to_insert[8] + "','" + data_to_insert[9] + "','" \
            + data_to_insert[10] + "','" + data_to_insert[11] + "','" + data_to_insert[12] + "','" \
            + data_to_insert[13] + "')"
        cursor.execute(query)

    conn.commit()

cursor.close()
conn.close()

    