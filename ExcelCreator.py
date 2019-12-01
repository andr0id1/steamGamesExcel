import xlsxwriter

from GameInfo import GameInfo

setup = [
    "name",
    "price argentina",
    "price chile",
    "price germany",
    "lowest price",
    "difference to ger",
    "chosen",
    "chosen name",
    "chosen price"
]

gameIds = ["1190460",
           "1091500",
           "474960",
           "804490",
           "850210",
           "1074420",
           "201810",
           "350080",
           "379720",
           "612880",
           "254700"]

euroclp = 0.001058
euroars = 0.015140

gameInfo = GameInfo(["ar", "cl", "de"])
col = 0
row = 0


def createGameRow(game_id, row, ):
    game_info = gameInfo.getGameInfo(game_id)
    cell_format1 = workbook.add_format()
    cell_format1.set_num_format(2)
    cell_format1.set_font_size(15)
    cell_format1.set_text_wrap()
    cell_format1.set_bold("true")

    cell_format2 = workbook.add_format()
    cell_format2.set_num_format(2)
    cell_format2.set_font_size(16)

    cell_format3 = workbook.add_format()
    cell_format3.set_num_format(10)
    cell_format3.set_font_size(16)
    cell_format3.set_text_wrap()
    worksheet.write(row, 0, game_info["name"], cell_format1)
    worksheet.write(row, 1, (game_info["ar"] * euroars) / 100, cell_format2)
    worksheet.write(row, 2, (game_info["cl"] * euroclp) / 100, cell_format2)
    worksheet.write(row, 3, game_info["de"] / 100, cell_format2)
    worksheet.write(row, 4, min((game_info["ar"] * euroars), (game_info["cl"] * euroclp),
                                game_info["de"]) / 100, cell_format2)
    worksheet.write(row, 5, (1 - (game_info["ar"] * euroars) / game_info["de"]), cell_format3)
    worksheet.write(row, 7, '=IF(G' + str(row + 1) + '="x",A' + str(row + 1) + ',"")', cell_format1)
    worksheet.write(row, 8, '=IF(G' + str(row + 1) + '="x",E' + str(row + 1) + ',"")', cell_format2)
    print(game_info["name"] + " written")


workbook = xlsxwriter.Workbook('hello.xlsx')
worksheet = workbook.add_worksheet('prices')
worksheet.set_column(0, 8, 12)

cell_format1 = workbook.add_format()
cell_format1.set_bg_color("#BEC0BF")
cell_format1.set_font_size(16)
cell_format1.set_text_wrap()
cell_format2 = workbook.add_format()
cell_format2.set_bg_color("#E6E8E8")
cell_format2.set_font_size(16)
cell_format2.set_num_format(2)
cell_format2.set_top(1)
for title in setup:
    worksheet.write(row, col, title, cell_format1)
    col += 1
row += 1

for gameId in gameIds:
    createGameRow(gameId, row)
    row += 1

worksheet.write(row, 0, "Total", cell_format2)
worksheet.write(row, 1, '=SUM(B2:B' + (str(row)) + ')', cell_format2)
worksheet.write(row, 2, '=SUM(C2:C' + (str(row)) + ')', cell_format2)
worksheet.write(row, 3, '=SUM(D2:D' + (str(row)) + ')', cell_format2)
worksheet.write(row, 4, '=SUM(E2:E' + (str(row)) + ')', cell_format2)
worksheet.write(row, 5, '', cell_format2)
worksheet.write(row, 6, '=COUNTA(G2:G' + (str(row)) + ')', cell_format2)
worksheet.write(row, 7, '', cell_format2)
worksheet.write(row, 8, '=SUM(I2:I' + (str(row)) + ')', cell_format2)

workbook.close()
