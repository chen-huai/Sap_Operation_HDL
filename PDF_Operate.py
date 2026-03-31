from io import StringIO
import os
import re
import pdfplumber

class PDF_Operate():
    def readPdf(inputFile):
        text = []
        with pdfplumber.open(inputFile) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text(x_tolerance=2)

                page_text_list = page_text.split("\n")

                text += page_text_list
        return text

    def saveAs(inputFile, outputFile):
        with open(inputFile, 'rb') as fp1:
            b1 = fp1.read()
        with open(outputFile, 'wb') as fp2:
            fp2.write(b1)

    def sanitize_filename(filename):
        """
        将文件名中的非法字符替换为下划线。

        Parameters:
        filename (str): 需要处理的文件名字符串。

        Returns:
        str: 替换非法字符后的文件名字符串。
        """
        # 定义 Windows 文件名中非法的字符
        invalid_chars = '<>:"/\\|?*'

        # 替换非法字符为下划线
        for char in invalid_chars:
            filename = filename.replace(char, '')

        return filename

# if __name__ == '__main__':
#     dirUrl = r"C:\Users\chen-fr\Desktop\临时文件\invoice"  # 文件夹目录
#     files = os.listdir(dirUrl)  # 得到文件夹下的所有文件名称
#     n = 1
#     invoiceNoStar = 4
#     orderNoStar = 7
#     msg = {}
#     pdfOperate = PDF_Operate
#     for each in files:
#         print(each)
#         fileUrl = '%s\\%s' % (dirUrl, each)
#         if os.path.isfile(fileUrl):
#             with open(fileUrl, 'rb') as my_pdf:
#                 print(n)
#                 fileCon = pdfOperate.readPdf(my_pdf)
#                 print(fileCon)
#                 fileNum = 0
#                 for fileCon[fileNum] in fileCon:
#                     if re.match('.*P. R. China', fileCon[fileNum]) or re.match('.*P.R. China',
#                                                                                fileCon[fileNum]) or re.match(
#                             'Pleasequotethisnumberonallinquiriesandpayments', fileCon[fileNum]) or  re.match(
#                             'Please quote this number on all inquiries and payments.', fileCon[fileNum]):
#                         if str(invoiceNoStar) in fileCon[fileNum + 1]:
#                             msg['Company Name'] = fileCon[fileNum + 2].replace(
#                                 'Please quote this number on all inquiries and payments.', '').replace(
#                                 'Invoice No.', '')
#                         else:
#                             msg['Company Name'] = fileCon[fileNum + 1].replace(
#                                 'Please quote this number on all inquiries and payments.', '').replace(
#                                 'Invoice No.', '')
#                     elif re.match('请在项目咨询或付款时提示此帐单号', fileCon[fileNum]):
#                         msg['Company Name'] = fileCon[fileNum + 2].replace(
#                             'Please quote this number on all inquiries and payments.', '').replace(
#                             'Invoice No.', '')
#                     elif re.match('%s\d{8}'%invoiceNoStar, fileCon[fileNum]):
#                         print(fileCon[fileNum], 22)
#                         msg['Invoice No'] = fileCon[fileNum]
#                     elif re.match('%s\d{8}'%orderNoStar, fileCon[fileNum]):
#                         print(fileCon[fileNum], 33)
#                         msg['Order No'] = fileCon[fileNum]
#                     elif re.match('\d{2}.\d{3}.\d{2}.\d{4,5}', fileCon[fileNum]):
#                         print(fileCon[fileNum], 44)
#                         msg['Project No'] = fileCon[fileNum]
#                     fileNum += 1
#                 n += 1
#                 outputFlie = msg['Project No'] + msg['Company Name'] + '.pdf'
#                 # outputFlie = msg['Invoice No'] + '-' + msg['Company Name'] + '.pdf'
#                 pdfOperate.saveAs(fileUrl,'%s\\test\\%s' % (dirUrl, outputFlie))
#
#         else:
#             print('无')





