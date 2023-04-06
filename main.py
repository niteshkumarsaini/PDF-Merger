import PyPDF2,win32com.client
speaker=win32com.client.Dispatch('SAPI.SpVoice')
speaker.Speak("First Put the Files in the Project Folder then.")
speaker.Speak("Select the Option below.")
try:
    while True:
        print("Enter 1 for Pdf Reading and Extracting Texts.. ")
        speaker.Speak("Enter 1 for Pdf Reading and Extracting Texts.")
        print("Enter 2 for Pdf Merging.")
        speaker.Speak("Enter 2 for Pdf Merging..")
        option=int(input(""))
        if(option==1):
            print("wait...")
            speaker.Speak("Enter the Pdf Name with Extension for Reading.")
            pdf=input("Enter Pdf Name. ")
            pdfRead=PyPDF2.PdfReader(pdf)
            speaker.Speak("Your Data is Below.")
        
            for i in pdfRead.pages:
                print(i.extract_text())
            break
        
        elif(option==2):
            print("Enter No. of Files you want to merge together.")
            speaker.Speak("Enter Number of Files you want to merge together.")
            totalFiles=int(input(""))
            fileList=[]
            speaker.Speak("Enter Files Name one by one with Extensions.")
            for i in range(totalFiles):
                file=input(f"Enter {i+1} file Name with Extension. ")
                fileList.append(file)
            mergedPdf=PyPDF2.PdfMerger()
            speaker.Speak("Congratulations your File is Ready.")
            for i in fileList:
                mergedPdf.append(i)
            mergedPdf.write('MergedPDF.pdf')
            break
        
        else:
            speaker.Speak("Select a Valid Option.")
except Exception as e:
    print(f"Something Went Wrong due to {e}")
    speaker.Speak("Something went Wrong. Please Restart the Program with valid Data.")
