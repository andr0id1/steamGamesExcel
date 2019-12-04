import xlsxwriter

from GameInfo import GameInfo


class ExcelCreator:
    global game_info
    global col
    global row
    global workbook
    global worksheet
    global euroclp
    global euroars
    global setup_colums
    global cell_format_first_row
    global cell_format_last_row
    global cell_format_gamename
    global cell_format_percentage
    global cell_format_price

    def __init__(self):
        self.workbook = xlsxwriter.Workbook('hello.xlsx')
        self.worksheet = self.workbook.add_worksheet('prices')
        self.game_info = GameInfo(["ar", "cl", "de"])
        self.col = 0
        self.row = 0
        self.euroclp = 0.001058
        self.euroars = 0.015148
        self.setup = [
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
        self.cell_format_first_row = self.workbook.add_format()
        self.cell_format_first_row.set_bg_color("#BEC0BF")
        self.cell_format_first_row.set_font_size(16)
        self.cell_format_first_row.set_text_wrap()

        self.cell_format_last_row = self.workbook.add_format()
        self.cell_format_last_row.set_bg_color("#E6E8E8")
        self.cell_format_last_row.set_font_size(16)
        self.cell_format_last_row.set_num_format(2)
        self.cell_format_last_row.set_top(1)

        self.cell_format_gamename = self.workbook.add_format()
        self.cell_format_gamename.set_num_format(2)
        self.cell_format_gamename.set_font_size(15)
        self.cell_format_gamename.set_text_wrap()
        self.cell_format_gamename.set_bold("true")

        self.cell_format_percentage = self.workbook.add_format()
        self.cell_format_percentage.set_num_format(10)
        self.cell_format_percentage.set_font_size(16)
        self.cell_format_percentage.set_text_wrap()

        self.cell_format_price = self.workbook.add_format()
        self.cell_format_price.set_num_format(2)
        self.cell_format_price.set_font_size(16)

    def create_first_row(self):
        for title in self.setup:
            self.worksheet.write(self.row, self.col, title, self.cell_format_first_row)
            self.col += 1
        self.row += 1

    def create_last_row(self):
        self.worksheet.write(self.row, 0, "Total", self.cell_format_last_row)
        self.worksheet.write(self.row, 1, '=SUM(B2:B' + (str(self.row)) + ')', self.cell_format_last_row)
        self.worksheet.write(self.row, 2, '=SUM(C2:C' + (str(self.row)) + ')', self.cell_format_last_row)
        self.worksheet.write(self.row, 3, '=SUM(D2:D' + (str(self.row)) + ')', self.cell_format_last_row)
        self.worksheet.write(self.row, 4, '=SUM(E2:E' + (str(self.row)) + ')', self.cell_format_last_row)
        self.worksheet.write(self.row, 5, '', self.cell_format_last_row)
        self.worksheet.write(self.row, 6, '=COUNTA(G2:G' + (str(self.row)) + ')', self.cell_format_last_row)
        self.worksheet.write(self.row, 7, '', self.cell_format_last_row)
        self.worksheet.write(self.row, 8, '=SUM(I2:I' + (str(self.row)) + ')', self.cell_format_last_row)

    def create_game_row(self, game_id):
        current_game_info = self.game_info.getGameInfo(game_id)
        self.worksheet.write(self.row, 0, current_game_info["name"], self.cell_format_gamename)
        self.worksheet.write(self.row, 1, (current_game_info["ar"] * self.euroars) / 100, self.cell_format_price)
        self.worksheet.write(self.row, 2, (current_game_info["cl"] * self.euroclp) / 100, self.cell_format_price)
        self.worksheet.write(self.row, 3, current_game_info["de"] / 100, self.cell_format_price)
        self.worksheet.write(self.row, 4, min((current_game_info["ar"] * self.euroars), (current_game_info["cl"] * self.euroclp),
                                         current_game_info["de"]) / 100, self.cell_format_price)
        self.worksheet.write(self.row, 5, (1 - (current_game_info["ar"] * self.euroars) / current_game_info["de"]), self.cell_format_percentage)
        self.worksheet.write(self.row, 7, '=IF(G' + str(self.row + 1) + '="x",A' + str(self.row + 1) + ',"")', self.cell_format_gamename)
        self.worksheet.write(self.row, 8, '=IF(G' + str(self.row + 1) + '="x",E' + str(self.row + 1) + ',"")', self.cell_format_price)
        print(current_game_info["name"] + " written")

    def create_excel(self, game_ids):
        self.worksheet.set_column(0, 8, 12)
        self.create_first_row()

        for gameId in game_ids:
            self.create_game_row(gameId)
            self.row += 1

        self.create_last_row()
        self.workbook.close()
