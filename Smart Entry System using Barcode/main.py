import cv2
import pyzbar.pyzbar as pyzbar
import datetime, keyboard
import pandas as pd
import xlsxwriter

workbook = xlsxwriter.Workbook('AllAboutPythonExcel.xlsx')
worksheet = workbook.add_worksheet()
cap = cv2.VideoCapture(0)
cap.set(3, 1000)  # Webcam Width & Height
cap.set(4, 600)

detector = cv2.QRCodeDetector()
# df = pd.DataFrame()
while cap.isOpened():

    success, image = cap.read()  # READ frames of video

    if not success:
        print('Skiping empty frame.')

    decodedObjects = pyzbar.decode(image)  # Decoding qr code data

    cv2.rectangle(image, (0, 0), (1000, 80), (245, 176, 66), -1)  # Background rectangle to display text
    worksheet.write(0, 0, "Branch")
    worksheet.write(0, 1, "Name")
    worksheet.write(0, 2, "Father's Name")
    worksheet.write(0, 3, "Address")
    # worksheet.write(0, 5, "Country")
    for obj in decodedObjects:
        data = obj.data.decode('utf-8')
        k = 0
        tmp = ''
        for element in data:
            if element == ',':
                worksheet.write(1, k, tmp)
                if k == 4:
                    break
                k = k + 1
                tmp = ''
            else:
                tmp = tmp + element
            # print(element, end=' ')
        # print("\n")
        # print(data)
        # df['Name'] = data
        # df.to_excel('result.xlsx', index=False)
        # worksheet.write(0, 0, "#")


        # for index, entry in enumerate(data):
        #     worksheet.write(index + 1, 0, str(index))
        #     worksheet.write(index + 1, 1, entry["name"])
        #     worksheet.write(index + 1, 2, entry["phone"])
        #     worksheet.write(index + 1, 3, entry["email"])
        #     worksheet.write(index + 1, 4, entry["address"])
        #     worksheet.write(index + 1, 5, entry["country"])

        workbook.close()


        cv2.putText(image, data, (20, 50), cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0, 0, 250), 2)  # Display data in qr code

        x, y, w, h = obj.rect  # get qr code coordinates to draw rectangle
        cv2.rectangle(image, (x, y), (x + w, y + h), (0, 250, 0), 2)  # Rectangle around qr code
        break

    cv2.imshow('QR CODE SCANNER', image)  # Show webcam feed
    if cv2.waitKey(5) & 0xFF == 27:
        break

cap.release()
cv2.destroyAllWindows()