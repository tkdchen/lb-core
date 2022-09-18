# -*- tab-width: 4; indent-tabs-mode: nil; py-indent-offset: 4 -*-
#
# This file is part of the LibreOffice project.
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
#

import officehelper

from com.sun.star.awt import Size
from com.sun.star.text.ControlCharacter import PARAGRAPH_BREAK
from com.sun.star.text.TextContentAnchorType import AS_CHARACTER


class ExampleDocument:

    def __init__(self):
        self.remote_context = officehelper.bootstrap()
        print("Connected to a running office ...")
        srv_mgr = self.remote_context.getServiceManager()
        desktop = srv_mgr.createInstanceWithContext("com.sun.star.frame.Desktop", self.remote_context)

        print("Opening an empty Writer document")
        doc_url = "private:factory/swriter"
        self.component = desktop.loadComponentFromURL(doc_url, "_blank", 0, tuple());

        self.text = self.component.getText()
        self.cursor = self.text.createTextCursor()

    def generate(self):
        self.insert_text("The first line in the newly created text document.\n")
        self.insert_text("Now we're in the second line\n")

        self.insert_table()
        self.insert_paragraph_break()

        print("Inserting colored Text")
        content = " This is a colored Text - blue with shadow\n"
        self.insert_text(content, CharColor=255, CharShadowed=True)

        self.insert_paragraph_break()
        self.insert_frame_with_text()
        self.insert_paragraph_break()
        self.insert_text(" That's all for now !!", CharColor=65536, CharShadowed=False)
        print("done")

    def insert_paragraph_break(self):
        self.text.insertControlCharacter(self.cursor, PARAGRAPH_BREAK, False)

    def insert_text(self, content, **props):
        for name, val in props.items():
            self.cursor.setPropertyValue(name, val)
        self.text.insertString(self.cursor, content, False)

    def insert_table(self):
        print("Inserting a text table")
        text_table = self.component.createInstance("com.sun.star.text.TextTable")
        text_table.initialize(4, 4)
        # Insert the table
        self.text.insertTextContent(self.cursor, text_table, False)
        # Get the first row
        rows = text_table.getRows()
        first_row = rows[0]

        # Set properties of the text table
        text_table.setPropertyValue("BackTransparent", False)
        text_table.setPropertyValue("BackColor", 13421823)
        # Set properties of the first row
        first_row.setPropertyValue("BackTransparent", False)
        first_row.setPropertyValue("BackColor", 6710932)

        print("Write text in the table headers")
        self.insert_into_cell("A1", "FirstColumn", text_table);
        self.insert_into_cell("B1", "SecondColumn", text_table);
        self.insert_into_cell("C1", "ThirdColumn", text_table);
        self.insert_into_cell("D1", "SUM", text_table);

        print("Insert something in the text table")
        data = (
            ("A2", 22.5, False),
            ("B2", 5615.3, False),
            ("C2", -2315.7, False),
            ("D2", "sum <A2:C2>", True),
            ("A3", 21.5, False),
            ("B3", 615.3, False),
            ("C3", -315.7, False),
            ("D3", "sum <A3:C3>", True),
            ("A4", 121.5, False),
            ("B4", -615.3, False),
            ("C4", 415.7, False),
            ("D4", "sum <A4:C4>", True),
        )
        for cell_name, value, is_formula in data:
            cell = text_table.getCellByName(cell_name)
            if is_formula:
                cell.setFormula(value)
            else:
                cell.setValue(value)

    def insert_frame_with_text(self):
        text_frame = self.component.createInstance("com.sun.star.text.TextFrame")
        frame_size = Size()
        frame_size.Height = 400
        frame_size.Width = 15000
        text_frame.setSize(frame_size)
        text_frame.setPropertyValue("AnchorType", AS_CHARACTER)
        self.text.insertTextContent(self.cursor, text_frame, False)

        print("Insert the text frame")
        frame_text = text_frame.getText()
        frame_cursor = frame_text.createTextCursor()
        text_frame.insertString(
            frame_cursor, "The first line in the newly created text frame.", False
        )
        text_frame.insertString(
            frame_cursor, "\nWith this second line the height of the frame raises.", False
        )

    def insert_into_cell(self, cell_name, text, text_table):
        """Insert text into specified table cell

        :param str cell_name: cell name.
        :param str text: the content to be inserted into cell.
        :param text_table: object which implements com.sun.star.text.XTextTable interface.
        """
        cell = text_table.getCellByName(cell_name)
        cursor = cell.createTextCursor()
        cursor.setPropertyValue("CharColor", 16777215)
        cell.setString(text)


if __name__ == "__main__":
    ExampleDocument().generate()

# vim: set shiftwidth=4 softtabstop=4 expandtab:
