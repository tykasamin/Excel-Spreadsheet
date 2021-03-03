import random
import xlsxwriter


def main():
    # tree id
    # of oranges
    # profit of tree
    # age of tree
    # type of tree

    workbook = xlsxwriter.Workbook("oranges.xlsx")

    bold_format = workbook.add_format({'bold': True})

    cell_format = workbook.add_format()
    cell_format.set_text_wrap()
    cell_format.set_align('top')
    cell_format.set_align('left=')

    moneyFormat = workbook.add_format({'num_format': '$#, ##0.00'})
    moneyRedFormat = workbook.add_format({'num_format': '$#, ##0.00'})
    moneyRedFormat.set_font_color('red')

    worksheet = workbook.add_worksheet("Oranges")

    worksheet.write('A1', 'Tree ID', bold_format)
    worksheet.write('B1', 'Number of Oranges', bold_format)
    worksheet.write('C1', 'Tree Height', bold_format)
    worksheet.write('D1', 'Tree Profit', bold_format)
    worksheet.write('E1', 'Tree Type', bold_format)

#start at row 2
    rowIndex = 2

    for row in range(200): #changing range can control the rows
        treeid = row + 1000
        numOranges = 20 + random.randint(50,100)
        typeOfTree = random.choice(['Navel', 'Valencia', 'Tangerines', 'Seville', 'Clementines'])
        treeProfit = random.random() * 1000
        heightOfTree = 100 + random.randint(25,50)

        worksheet.write('A' + str(rowIndex), treeid, cell_format)
        worksheet.write('B' + str(rowIndex), numOranges, cell_format)
        worksheet.write('C' + str(rowIndex), heightOfTree, cell_format)

        if treeProfit < 100.0:
            worksheet.write('D' + str(rowIndex), treeProfit, moneyRedFormat)
        else:
            worksheet.write('D' + str(rowIndex), treeProfit, moneyFormat)

        worksheet.write('E' + str(rowIndex), typeOfTree, cell_format)

        rowIndex += 1

        print(treeid, numOranges, typeOfTree, treeProfit, heightOfTree)

    worksheet.set_column(4,4,width=25) #make column Tree Type wider
    workbook.close() #important!

if __name__ == "__main__":
    main()

