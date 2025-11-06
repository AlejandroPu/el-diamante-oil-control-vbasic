Attribute VB_Name = "a_datosCopec"
' no funciona
Sub accesoCopec()
    Dim IE As New SHDocVw.InternetExplorer
    IE.Visible = True
    IE.Navigate "http://www.copec.cl/wp7/wps/myportal/!ut/p/b1/jY9LDoIwEEDPwglmaAstywZpS2wkSqrSjWFhDIbPxnh-gbhxYXV2k7yXNwMeGppS5CwTCGfwY_vsbu2jm8a2X3afXmri4o0yBLU9CiS0snVcuVhv2Qw0M4BfRuKnL4zjSKzitizyBXj7uZaGcYsorE6wlMYdsj2lKOl__UDgh38CvyKhC1Yg9GIwUiWwM9NwhcH3SmXlnckoegEAJ5go/dl4/d5/L2dJQSEvUUt3QS80SmtFL1o2X1MyVTFERkgyMEdMVjgwMjNPTFMxT1UxRzI0/"
    
    Do While IE.ReadyState <> READYSTATE_COMPLETE
    Loop
    
    IE.Document.forms(0).elements("username").Value = "{erased in this script for privacy}"
    IE.Document.forms(0).elements("password").Value = "{erased in this script for privacy}"
    'IE.Document.forms(0).submit
    
End Sub
