# -*- tab-width: 4; indent-tabs-mode: nil; py-indent-offset: 4 -*-
#
# This file is part of the LibreOffice project.
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
#

import officehelper


def main():
    remote_context = officehelper.bootstrap()
    print("Connected to a running office ...")
    srv_mgr = remote_context.getServiceManager()
    desktop = srv_mgr.createInstanceWithContext("com.sun.star.frame.Desktop", remote_context)

    print("Opening an empty Writer document")
    doc_url = "private:factory/swriter"
    component = desktop.loadComponentFromURL(doc_url, "_blank", 0, tuple());

    text = component.getText()
    text.setString("Please select something in this text and press then \"return\" in the shell "
                    "where you have started the example.\n")

    # Returned object supports service com.sun.star.text.TextDocumentView and com.sun.star.view.OfficeDocumentView
    # Both of them implements interface com::sun::star::view::XViewSettingsSupplier
    obj = component.getCurrentController()
    obj.getViewSettings().setPropertyValue("ZoomType", 0)

    print()
    input("Please select something in the test document and press "
          "then \"return\" to continues the example ... ")

    frame = desktop.getCurrentFrame()
    selection = frame.getController().getSelection()

    if selection.supportsService("com.sun.star.text.TextRanges"):
        for selected in selection:
            print("You have selected a text range:", f'"{selected.getString()}".')

    if selection.supportsService("com.sun.star.text.TextGraphicObject"):
        print( "You have selected a graphics." )

    if selection.supportsService("com.sun.star.text.TexttableCursor"):
        print( "You have selected a text table." )

    component.dispose()


if __name__ == "__main__":
    main()

# vim: set shiftwidth=4 softtabstop=4 expandtab:
