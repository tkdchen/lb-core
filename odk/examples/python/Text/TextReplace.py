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

    def do_replace(self):
        self.create_example_data()

        british_words = ["colour", "neighbour", "centre", "behaviour", "metre", "through"]
        us_words = ["color", "neighbor", "center", "behavior", "meter", "thru"]

        replace_descriptor = self.component.createReplaceDescriptor()
        print("Change all occurrences of ...")
        for british_word, us_word in zip(british_words, us_words):
            replace_descriptor.setSearchString(british_word)
            replace_descriptor.setReplaceString(us_word)
            # Replace all words
            replaced_cnt = self.component.replaceAll(replace_descriptor)
            if replaced_cnt > 0:
                print("Replaced", british_word, "with", us_word)

        print("Done")

    def create_example_data(self):
        text = self.component.getText()
        cursor = text.createTextCursor()
        text.insertString(cursor, "He nervously looked all around. Suddenly he saw his ", False);

        text.insertString(cursor, "neighbour ", True);
        cursor.setPropertyValue("CharColor", 255)  # Set the word blue

        cursor.gotoEnd(False)  # Go to last character
        cursor.setPropertyValue("CharColor", 0)
        content = (
            "in the alley. Like lightning he darted off to the left and disappeared between the "
            "two warehouses almost falling over the trash can lying in the "
        )
        text.insertString(cursor, content, False);

        text.insertString(cursor, "centre ", True);
        cursor.setPropertyValue("CharColor", 255)  # Set the word blue

        cursor.gotoEnd(False)  # Go to last character
        cursor.setPropertyValue("CharColor", 0)
        text.insertString(cursor, "of the sidewalk.", False);

        text.insertControlCharacter(cursor, PARAGRAPH_BREAK, False)
        content = (
            "He tried to nervously tap his way along in the inky darkness and suddenly stiffened: "
            "it was a dead-end, he would have to go back the way he had come."
        )
        text.insertString(cursor, content, False);
        cursor.gotoStart(False)


if __name__ == "__main__":
    ExampleDocument().do_replace()

# vim: set shiftwidth=4 softtabstop=4 expandtab:
