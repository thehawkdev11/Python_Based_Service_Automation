from tkinter import *
from PIL import Image,ImageTk
from tkinter import messagebox
import num2words
import pyshorteners
from time import sleep
import gspread
from oauth2client.service_account import ServiceAccountCredentials

import pywhatkit

import DateTime

scope = ['https://www.googleapis.com/auth/drive']

credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)

client = gspread.authorize(credentials)

sheet = client.open("Testing")

wc = sheet.worksheet("Sheet1")

class Service:
    def __init__(self,root):
        self.root = root
        self.root.title("Multisyns customer Service")
        self.root.geometry("1350x700+0+0")
        self.root.configure(background='lightgrey')

        # Service Frame
        frame = Frame(self.root,bg="white")
        frame.place(x=50,y = 100,width=1250,height=580)

        title = Label(frame, text="MULTISYNS CUSTOMER SERVICE FORM",font=("times new roman",15,"bold"),bg="white").place(x=420, y =10)
        # Customer Info
        customers_name = Label(frame, text="Customer's Name",font=("times new roman",13,"bold"),bg="white").place(x=220, y =50)
        self.txt_customers_name = Entry(frame,font=("times new roman",13,),bg="white")
        self.txt_customers_name.place(x=220, y=80,width=250,height=30)

        customers_number = Label(frame, text="Customer's Number", font=("times new roman", 13, "bold"), bg="white").place(x=500, y=50)
        self.txt_customers_number = Entry(frame, font=("times new roman", 13,), bg="white")
        self.txt_customers_number.place(x=500, y=80, width=250,height=30)

        customers_email = Label(frame, text="Customer's Email", font=("times new roman", 13, "bold"), bg="white").place(x=780, y=50)
        self.txt_customers_email = Entry(frame, font=("times new roman", 13,), bg="white")
        self.txt_customers_email.place(x=780, y=80, width=250,height=30)
        # Order Details

        # Row -1
        media_1 = Label(frame, text="Media",font=("times new roman",13,"bold"),bg="white").place(x=280, y =120)
        self.txt_media_1 = Entry(frame, font=("times new roman", 13,), bg="white")
        self.txt_media_1.place(x=280, y=150, width=50,height=30)

        job_1 = Label(frame, text="Job Title", font=("times new roman", 13, "bold"), bg="white").place(x=450, y=120)
        self.txt_job_1 = Entry(frame, font=("times new roman", 13,), bg="white")
        self.txt_job_1.place(x=350, y=150, width=250, height=30)

        size_1 = Label(frame, text="Size", font=("times new roman", 13, "bold"), bg="white").place(x=630, y=120)
        self.txt_size_1 = Entry(frame, font=("times new roman", 13,), bg="white")
        self.txt_size_1.place(x=630, y=150, width=50, height=30)

        qty_1 = Label(frame, text="Qty", font=("times new roman", 13, "bold"), bg="white").place(x=700, y=120)
        self.txt_qty_1 = Entry(frame, font=("times new roman", 13,), bg="white")
        self.txt_qty_1.place(x=700, y=150, width=50, height=30)

        sqft_1 = Label(frame, text="SQFT", font=("times new roman", 13, "bold"), bg="white").place(x=770, y=120)
        self.txt_sqft_1 = Entry(frame, font=("times new roman", 13,), bg="white")
        self.txt_sqft_1.place(x=770, y=150, width=50, height=30)

        amount_1 = Label(frame, text="Amount", font=("times new roman", 13, "bold"), bg="white").place(x=840, y=120)
        self.txt_amount_1 = Entry(frame, font=("times new roman", 13,), bg="white")
        self.txt_amount_1.place(x=840, y=150, width=100, height=30)

        # Row -2
        self.txt_media_2 = Entry(frame, font=("times new roman", 13,), bg="white")
        self.txt_media_2.place(x=280, y=200, width=50, height=30)
        self.txt_job_2 = Entry(frame, font=("times new roman", 13,), bg="white")
        self.txt_job_2.place(x=350, y=200, width=250, height=30)
        self.txt_size_2 = Entry(frame, font=("times new roman", 13,), bg="white")
        self.txt_size_2.place(x=630, y=200, width=50, height=30)
        self.txt_qty_2 = Entry(frame, font=("times new roman", 13,), bg="white")
        self.txt_qty_2.place(x=700, y=200, width=50, height=30)
        self.txt_sqft_2 = Entry(frame, font=("times new roman", 13,), bg="white")
        self.txt_sqft_2.place(x=770, y=200, width=50, height=30)
        self.txt_amount_2 = Entry(frame, font=("times new roman", 13,), bg="white")
        self.txt_amount_2.place(x=840, y=200, width=100, height=30)
        # Row - 3
        self.txt_media_3 = Entry(frame, font=("times new roman", 13,), bg="white")
        self.txt_media_3.place(x=280, y=250, width=50, height=30)
        self.txt_job_3 = Entry(frame, font=("times new roman", 13,), bg="white")
        self.txt_job_3.place(x=350, y=250, width=250, height=30)
        self.txt_size_3 = Entry(frame, font=("times new roman", 13,), bg="white")
        self.txt_size_3.place(x=630, y=250, width=50, height=30)
        self.txt_qty_3 = Entry(frame, font=("times new roman", 13,), bg="white")
        self.txt_qty_3.place(x=700, y=250, width=50, height=30)
        self.txt_sqft_3 = Entry(frame, font=("times new roman", 13,), bg="white")
        self.txt_sqft_3.place(x=770, y=250, width=50, height=30)
        self.txt_amount_3 = Entry(frame, font=("times new roman", 13,), bg="white")
        self.txt_amount_3.place(x=840, y=250, width=100, height=30)
        # row - 4
        self.txt_media_4 = Entry(frame, font=("times new roman", 13,), bg="white")
        self.txt_media_4.place(x=280, y=300, width=50, height=30)
        self.txt_job_4 = Entry(frame, font=("times new roman", 13,), bg="white")
        self.txt_job_4.place(x=350, y=300, width=250, height=30)
        self.txt_size_4 = Entry(frame, font=("times new roman", 13,), bg="white")
        self.txt_size_4.place(x=630, y=300, width=50, height=30)
        self.txt_qty_4 = Entry(frame, font=("times new roman", 13,), bg="white")
        self.txt_qty_4.place(x=700, y=300, width=50, height=30)
        self.txt_sqft_4 = Entry(frame, font=("times new roman", 13,), bg="white")
        self.txt_sqft_4.place(x=770, y=300, width=50, height=30)
        self.txt_amount_4 = Entry(frame, font=("times new roman", 13,), bg="white")
        self.txt_amount_4.place(x=840, y=300, width=100, height=30)
        # submit button

        btn = Button(frame,text="Submit",bd=0,cursor="hand2",command=self.ServiceData,bg="green",font=("times new roman",12,"bold")).place(x=600,y=340,height=30,width="80")

        ################################ Search Section ##########################
        invoice = Label(frame, text="Invoice No: ", font=("times new roman", 13, "bold"), bg="white").place(x=400, y=400)
        self.txt_invoice_number = Entry(frame, font=("times new roman", 13,), bg="white")
        self.txt_invoice_number.place(x=510, y=400, width=250, height=30)

        search_btn = Button(frame,text="Search",bd=0,cursor="hand2",bg="green",font=("times new roman",12,"bold")).place(x=810,y=400,height=30,width="80")



    def ServiceData(self):
        if self.txt_customers_name.get()=="" or self.txt_customers_number.get()=="" or self.txt_amount_1.get()=="" or self.txt_media_1.get()=="" or self.txt_job_1.get()=="" or self.txt_qty_1.get()=="" or self.txt_size_1.get()=="" or self.txt_sqft_1.get()=="":
            messagebox.showerror("Error","Required Fields Are Missing",parent=self.root)
        index1 = len(wc.get_all_records()) + 2
        index2 = index1 + 1
        index3 = index1 + 2
        index4 = index1 + 3
        time = DateTime.DateTime()
        time = str(time)
        time = time.split(' ')
        Date = time[0]
        time = time[1]
        inv_no = "INV_000" + str(index1)
        total_amount = int(self.txt_amount_1.get()) + int(self.txt_amount_2.get()) + int(self.txt_amount_3.get()) + int(self.txt_amount_4.get())
        payment_status = 'Unpaid'
        row1 = [time,self.txt_customers_name.get(),self.txt_customers_number.get(),self.txt_customers_email.get(),self.txt_media_1.get(),self.txt_job_1.get(),self.txt_size_1.get(),self.txt_qty_1.get(),self.txt_sqft_1.get(),self.txt_amount_1.get(),total_amount,inv_no,payment_status]
        row2 = [" "," "," "," ",self.txt_media_2.get(),self.txt_job_2.get(),self.txt_size_2.get(),self.txt_qty_2.get(),self.txt_sqft_2.get(),self.txt_amount_2.get()]
        row3 = [" ", " ", " ", " ", self.txt_media_3.get(), self.txt_job_3.get(), self.txt_size_3.get(),self.txt_qty_3.get(),self.txt_sqft_3.get(),self.txt_amount_3.get()]
        row4 = [" ", " ", " ", " ", self.txt_media_4.get(), self.txt_job_4.get(), self.txt_size_4.get(),self.txt_qty_4.get(),self.txt_sqft_4.get(),self.txt_amount_4.get()]

        print(len(wc.get_all_records()))
        print(index1)
        print(index2)
        print(index3)
        print(index4)
        wc.insert_row(row1,index1)
        wc.insert_row(row2, index2)
        wc.insert_row(row3,index3)
        wc.insert_row(row4, index4)


        ########################### PAYMENT LINK GENERATION #########################

        base_url = 'https://www.upiqrcode.com/upi-link/?apikey=iyouvg&seckey=arbazaaa&vpa=multisynsads@okaxis&payee=multisyns&billno='+inv_no+'&amount='

        payment_url = base_url + str(total_amount)
        s = pyshorteners.Shortener()
        payment_url = s.tinyurl.short(payment_url)


        ############################# PDF GENERATION PART ###########################

        # ###################################
        # Content
        fileName = inv_no+".pdf"
        documentTitle = inv_no

        image = 'Multisyns logo-page-001.jpg'

        # ###################################
        # 0) Create document
        from reportlab.pdfgen import canvas

        pdf = canvas.Canvas(fileName)
        pdf.setTitle(documentTitle)

        pdf.setFont("Times-Bold", 22)
        pdf.drawString(40, 760, "A.K Unity's Multisyns")
        pdf.setFont('Times-Roman', 12)
        pdf.drawString(40, 740, 'Khilwath, Charminar,Hyderabad')
        pdf.setFont('Times-Roman', 12)
        pdf.drawString(40, 720, 'Phone: 9966649435')

        ### Image ####

        pdf.drawInlineImage(image, 360, 720, width=200, height=50)

        # ###################################
        # 2) Sub Title
        # RGB - Red Green and Blue
        pdf.setFillColorRGB(255, 0, 0)
        pdf.setFont("Times-Bold", 20)
        pdf.drawCentredString(300, 680, 'Tax Invoice')

        ############# Customer Details ##########
        pdf.setFillColorRGB(0, 0, 0)
        pdf.setFont("Times-Bold", 14)
        pdf.drawString(40, 650, 'Bill to: ')
        pdf.setFont("Times-Roman", 12)
        pdf.drawString(40, 630, 'c/o '+self.txt_customers_name.get())
        pdf.drawString(40, 610, 'Conatct No: ')
        pdf.drawString(120, 610, self.txt_customers_number.get())
        pdf.setFont("Times-Bold", 12)
        pdf.drawString(400, 650, 'Invoice no:')
        pdf.drawString(400, 630, 'Date:')
        pdf.setFont("Times-Roman", 12)
        pdf.drawString(500, 650, inv_no)
        pdf.drawString(500, 630, str(Date))

        # ###################################
        # 3) Draw a line
        pdf.line(40, 600, 560, 600)
        pdf.drawString(50, 585, '#')
        pdf.drawString(80, 585, 'Item Name')
        pdf.drawString(300, 585, 'Media')
        pdf.drawString(350, 585, 'Qty')
        pdf.drawString(400, 585, 'SQFT')
        pdf.drawString(500, 585, 'Amount')
        pdf.line(40, 575, 560, 575)

        ########### Entry 1 ###############
        pdf.drawString(50, 555, '1')
        pdf.drawString(80, 555, self.txt_job_1.get())
        pdf.drawString(300, 555, self.txt_media_1.get())
        pdf.drawString(350, 555, self.txt_qty_1.get())
        pdf.drawString(400, 555, self.txt_sqft_1.get())
        pdf.drawString(500, 555, self.txt_amount_1.get())

        ########## Entry - 2 ##############

        pdf.drawString(50, 535, '2')
        pdf.drawString(80, 535, self.txt_job_2.get())
        pdf.drawString(300, 535, self.txt_media_2.get())
        pdf.drawString(350, 535, self.txt_qty_2.get())
        pdf.drawString(400, 535, self.txt_sqft_2.get())
        pdf.drawString(500, 535, self.txt_amount_2.get())

        ########## Entry - 3 ##############

        pdf.drawString(50, 515, '3')
        pdf.drawString(80, 515, self.txt_job_3.get())
        pdf.drawString(300, 515, self.txt_media_3.get())
        pdf.drawString(350, 515, self.txt_qty_3.get())
        pdf.drawString(400, 515, self.txt_sqft_3.get())
        pdf.drawString(500, 515, self.txt_amount_3.get())

        ########## Entry - 4 ##############

        pdf.drawString(50, 495, '4')
        pdf.drawString(80, 495, self.txt_job_4.get())
        pdf.drawString(300, 495, self.txt_media_4.get())
        pdf.drawString(350, 495, self.txt_qty_4.get())
        pdf.drawString(400, 495, self.txt_sqft_4.get())
        pdf.drawString(500, 495, self.txt_amount_4.get())

        total_items = int(self.txt_qty_1.get())+int(self.txt_qty_2.get())+int(self.txt_qty_3.get())+int(self.txt_qty_4.get())


        pdf.line(40, 475, 560, 475)
        pdf.setFont('Times-Bold', 12)
        pdf.drawString(80, 460, 'Total')
        pdf.drawString(350, 460, str(total_items))
        pdf.drawString(500, 460, str(total_amount))
        pdf.line(40, 450, 560, 450)

        pdf.setFont('Times-Roman', 12)
        pdf.drawString(350, 430, 'Sub Total: ')
        pdf.drawString(350, 410, 'Additional Amount: ')
        pdf.setFont('Times-Bold', 13)
        pdf.drawString(350, 390, 'Total: ')
        pdf.setFont('Times-Roman', 12)
        pdf.drawString(350, 370, 'Received: ')
        pdf.drawString(350, 350, 'Net Balance: ')
        pdf.drawString(350, 330, 'Total Saving: ')

        pdf.setFont('Times-Roman', 12)
        pdf.drawString(500, 430, str(total_amount))
        pdf.drawString(500, 410, '0.00')
        pdf.setFont('Times-Bold', 13)
        pdf.drawString(500, 390, str(total_amount))
        pdf.setFont('Times-Roman', 12)
        pdf.drawString(500, 370, '0.00')
        pdf.drawString(500, 350, str(total_amount))
        pdf.drawString(500, 330, '300.00')

        # amount_in_words = num2words(total_amount)
        # amount_in_words = amount_in_words.split('-')
        # amount_in_words = " ".join(amount_in_words)
        # amount_in_words = amount_in_words.title()
        pdf.setFont('Times-Bold', 15)
        pdf.drawString(40, 430, 'Invoice Amount In Words')
        pdf.setFont('Times-Roman', 11)
        pdf.drawString(40, 410, "amount_in_words")

        pdf.save()
        print(total_amount)
        sleep(5)

        ########################## PDF UPLOADING TO DRIVE ##########################

        import json
        import requests
        name_of_file = inv_no+".pdf"
        path ="./"+name_of_file
        headers = {
            "Authorization": "Bearer ya29.a0AfH6SMBmHMTbsg2OUK1CjYxZg2jUbY8HTiV8EML7Q6zL-meUR1ZjvbVjSwazaPnEYbpZx0BerdoyW2IYuZxF6FiG84rtE-ZOKCsRlNw6w8Xji3flG7og-krRFKLqWdAzOfY7OrUfBRoWOvL8S5b1iFbLsYWHKBxLYjBXvBMGUMs"}
        para = {
            "name": name_of_file,
            "parents": ["1TxJduhCtP_iP33ycTDeIE0fj5DVS6F5C"]
        }
        files = {
            'data': ('metadata', json.dumps(para), 'application/json; charset=UTF-8'),
            'file': open(path, "rb")
        }
        r = requests.post(
            "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart",
            headers=headers,
            files=files
        )
        print(r.text)

        gdrive_link = 'https://7twaotzse5eo6jfdoexvsw-on.drv.tw/Invoices/'+inv_no+".pdf"



        ######################### SENDING WHATSAPP MESSAGE ##########################

        message_body = "Hi," + self.txt_customers_name.get() +"\nThank you for choosing multisyns for your job," + "\nyour *Invoice No : "+inv_no+"*"+"\n Invoice Link: "+gdrive_link+"\n Payment Link: "+payment_url
        time_w = DateTime.DateTime()

        time_w = str(time_w)
        time_w = time_w.split(' ')
        time_w = time_w[1]
        hours_w = int(time_w[0:2])

        Minutes_w = int(time_w[3:5])
        Minutes_w = Minutes_w + 1.2


        pywhatkit.sendwhatmsg("+91"+self.txt_customers_number.get(), message_body, hours_w, Minutes_w)



root = Tk()
obj = Service(root)
root.mainloop()
