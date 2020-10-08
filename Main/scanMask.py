import pandas as pd
import xlrd, time, re, csv, Main
from Main import CreditCardValidation as credit
import docx2txt, fileinput, docx

ts = time.time()
timestr = time.strftime("%Y%m%d-%H%M%s")


# start point
def main():
    data = pd.read_csv(Main.location_of_excel, dtype=str, delimiter="|")

    # for docx
    my_text = docx2txt.process(Main.location_of_excel)
    with open("MaskedFile/doc_text.txt", "w+") as writer:
        writer.write(my_text)
    writer.close()
    data = pd.read_csv("MaskedFile/doc_text.txt", dtype=str, delimiter="|")

    header_names = list(data.head(0))
    print(header_names)
    mainDataFrame = pd.DataFrame(data, dtype=str)
    mainDataFrame = mainDataFrame.astype(str)
    arrofUnmasked = []
    binNumbers = []
    noOfRowsMarked = 0

    for header in header_names:
        header = str(header)
        print("column begins:"+header+"--------------")
        newdata = pd.DataFrame(data, columns=[header])

        if header in Main.headersToEliminate:
            continue
        else:
            for ind, row in newdata.iterrows():
                columnsentence = row[header]
                re_list = [
                    r"(?=(54\d{14}))",
                    r"(?=(55\d{14}))",
                    r"(?=(4\d{15}))",
                    r"(?=(51\d{14}))",
                    r"(?=(34\d{13}))",
                    r"(?=(37\d{13}))",
                    r"(?=(622\d{13}))",
                    r"(?=(622\d{14}))",
                    r"(?=(622\d{16}))",
                    r"(?=(300\d{11}))",
                    r"(?=(305\d{11}))",
                    r"(?=(36\d{12}))",
                    r"(?=(6\d{15}))",
                    r"(?=(3528\d{12}))",
                    r"(?=(3589\d{12}))",
                    r"(?=(5018\d{8}))",
                    r"(?=(5018\d{9}))",
                    r"(?=(5018\d{10}))",
                    r"(?=(5018\d{11}))",
                    r"(?=(5018\d{12}))",
                    r"(?=(5018\d{13}))",
                    r"(?=(5018\d{14}))",
                    r"(?=(5018\d{15}))",
                    r"(?=(5020\d{8}))",
                    r"(?=(5020\d{9}))",
                    r"(?=(5020\d{10}))",
                    r"(?=(5020\d{11}))",
                    r"(?=(5020\d{12}))",
                    r"(?=(5020\d{13}))",
                    r"(?=(5020\d{14}))",
                    r"(?=(5020\d{15}))",
                ]
                matches = []
                strcolumnsentence = str(columnsentence)
                for r in re_list:
                    matches += re.findall(r, strcolumnsentence)
                    # print(matches)
                for match in matches:
                    value, maskedCard, binNumber = credit.CreditCardValidation(match, strcolumnsentence).startValidation()
                    if value:
                        # if len(binNumber) > 1:
                        #     binNumber.append(binNumber)
                        print(maskedCard)
                        if match not in columnsentence:
                            continue
                        indextoSplit = columnsentence.find(match)
                        columnsentence = columnsentence[:indextoSplit] + maskedCard + columnsentence[indextoSplit+len(maskedCard):]
                        noOfRowsMarked += 1
                        mainDataFrame.at[ind, str(header)] = columnsentence
                mainDataFrame.at[ind, str(header)] = columnsentence

    mainDataFrame.to_csv("MaskedFile/MaskedFile"+timestr+".txt", index=None, sep='|', quoting=csv.QUOTE_NONE)
    #for docx
    mainDataFrame.to_csv("MaskedFile/temp_data.txt", index=None, sep='|', quoting=csv.QUOTE_NONE)
    print("removing |")
    for line in fileinput.input("MaskedFile/temp_data.txt", inplace=True):
        print(line.replace("|\n", ""), end="")
    document = docx.Document()
    myfile = open("MaskedFile/temp_data.txt").read()
    myfile = re.sub(r'[^\x00-\x7F]+|\x0c', ' ',myfile)
    document.add_paragraph(myfile)
    name = "MaskedFile/MaskedFile"+ timestr+ ".docx"
    document.save(name)

    print("-----Finished Successfully-----")
    input("Press ENTER TO EXIT....")


if __name__ == '__main__':
    main()

