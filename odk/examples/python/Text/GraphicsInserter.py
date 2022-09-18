# -*- tab-width: 4; indent-tabs-mode: nil; py-indent-offset: 4 -*-
#
# This file is part of the LibreOffice project.
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
#

import argparse
from os.path import isfile
from posixpath import abspath, realpath

import officehelper
from com.sun.star.beans import PropertyValue
from com.sun.star.text.TextContentAnchorType import AT_PARAGRAPH


def insert_graphic(filename):
    remote_context = officehelper.bootstrap()
    srv_mgr = remote_context.getServiceManager()
    desktop = srv_mgr.createInstanceWithContext("com.sun.star.frame.Desktop", remote_context)

    doc_url = "private:factory/swriter"
    component = desktop.loadComponentFromURL(doc_url, "_blank", 0, tuple());

    text = component.getText()
    cursor = text.createTextCursor()

    print("inserting graphic")
    graphic = component.createInstance("com.sun.star.text.TextGraphicObject")
    text.insertTextContent(cursor, graphic, True)

    print("prepare graphic provider for the image to be displayed")
    graphic_url = f"file://{filename}".replace("\\", "/")
    print("insert graphic:", graphic_url)
    graphic_provider = srv_mgr.createInstanceWithContext(
        "com.sun.star.graphic.GraphicProvider", remote_context
    )
    loaded_graphic = graphic_provider.queryGraphic(
        (PropertyValue(Name="URL", Value=graphic_url),)
    )

    # Set properties for the inserted graphic
    graphic.setPropertyValue("AnchorType", AT_PARAGRAPH)
    # Setting the graphic url
    graphic.setPropertyValue("Graphic", loaded_graphic)
    # Setting the horizontal position
    graphic.setPropertyValue("HoriOrientPosition", 5500)
    # Setting the vertical position
    graphic.setPropertyValue("VertOrientPosition", 4200)
    # Setting the width
    graphic.setPropertyValue( "Width", 4400)
    # Setting the height
    graphic.setPropertyValue("Height", 4000)


def is_file(value):
    if not isfile(value):
        raise argparse.ArgumentTypeError(f"File {value} is not an image file.")
    return value


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("image", type=is_file, help="Path to an image file.")
    args = parser.parse_args()
    insert_graphic(args.image)


if __name__ == "__main__":
    main()

# vim: set shiftwidth=4 softtabstop=4 expandtab:
