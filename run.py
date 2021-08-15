import docx
import os
path = os.getcwd()
num = 0


def getText(filename):
    doc = docx.Document(filename)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    return '\n　　'.join(fullText)
        

def new_novel_txt():
    if not os.path.exists("novel_txt"):
        os.mkdir("novel_txt")
        
for file_name in os.listdir(path):
    new_file_name = file_name.split(".")
    new_novel_txt()
    num += 1
    if new_file_name[-1] == "docx":
        try:
            file = getText(file_name)
            print("正在处理第{}本:{}".format(num, file_name))
        except:
            print("第{}本处理失败:文件名{}".format(num, file_name))
            with open("处理失败.txt", "a") as f:
                f.write('\n' + new_file_name[0])
            continue
        open(os.path.join("novel_txt", f"{new_file_name[0]}.txt"), "w").write(file)
    else:
        print(file_name,"不是.docx文件，不做处理")
        continue
    
    