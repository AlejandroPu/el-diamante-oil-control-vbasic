Attribute VB_Name = "zTrabajoConBrowser"
Option Explicit

'creado con ayuda de https://youtu.be/dShR33CdlY8

Sub BrowseToSite()
    Dim IE As Object
    Set IE = CreateObject("InternetExplorer.Application")
    IE.Visible = True
    IE.Navigate "https://www.copec.cl/wp7/wps/myportal"
End Sub
' This code is the same, but next to add the Tool->refers->Microsoft Internet Controls
Sub BrowseToSite2()
    Dim IE As New SHDocVw.InternetExplorer
    IE.Visible = True
    IE.Navigate "https://www.copec.cl/wp7/wps/myportal"
End Sub

Sub ReadyStateBrowsing()
    Dim IE As New SHDocVw.InternetExplorer
    IE.Visible = True
    IE.Navigate "www.copec.cl/wp7/wps/myportal"
    
    Do While IE.ReadyState <> READYSTATE_COMPLETE
        'this loop could have not any of the following instructions code
        Application.Wait Now + TimeValue("00:00:01")
        DoEvents
    Loop
    Debug.Print IE.LocationName, IE.LocationURL
End Sub
Sub GettingElements()
    Dim IE As New SHDocVw.InternetExplorer
    IE.Visible = True
    IE.Navigate "www.copec.cl/wp7/wps/myportal"
    
    Do While IE.ReadyState <> READYSTATE_COMPLETE
        'this loop could have not any of the following instructions code
        Application.Wait Now + TimeValue("00:00:01")
        DoEvents
    Loop
    
    IE.Document.forms(0).elements("username").Value = "{erased in this script for privacy}"
    IE.Document.forms(0).elements("password").Value = "{erased in this script for privacy}"
    'IE.Document.forms(0).submit
    
End Sub
' Here I added Tool->refers->Microsoft HTML Object Library
' This code wasn't tested, is only for reference
Sub GetHTMLDocument()
    Dim IE As New SHDocVw.InternetExplorer
    Dim HTMLDoc As MSHTML.HTMLDocument
'    Dim HTMLInput As MSHTML.IHTMLElement
    Dim HTMLButtons As MSHTML.IHTMLElementCollection
    Dim HTMLButton As MSHTML.IHTMLElement
    
    IE.Visible = True
    IE.Navigate "http://www.copec.cl/wp7/wps/myportal/!ut/p/b1/jY9LDoIwEEDPwglmaAstywZpS2wkSqrSjWFhDIbPxnh-gbhxYXV2k7yXNwMeGppS5CwTCGfwY_vsbu2jm8a2X3afXmri4o0yBLU9CiS0snVcuVhv2Qw0M4BfRuKnL4zjSKzitizyBXj7uZaGcYsorE6wlMYdsj2lKOl__UDgh38CvyKhC1Yg9GIwUiWwM9NwhcH3SmXlnckoegEAJ5go/dl4/d5/L2dJQSEvUUt3QS80SmtFL1o2X1MyVTFERkgyMEdMVjgwMjNPTFMxT1UxRzI0/"
    
    Do While IE.ReadyState <> READYSTATE_COMPLETE
    Loop
    
    Set HTMLDoc = IE.Document
'    Set HTMLInput = HTMLDoc.GetElementById("id_del_elemento")
'    HTMLInput.Value = "palabra_a_ingresas"
    
    Set HTMLButtons = HTMLDoc.getElementsByTagName("input") 'Theorically her say "button"

    For Each HTMLButton In HTMLButtons
        Debug.Print HTMLButton.className, HTMLButton.tagName, HTMLButton.ID, HTMLButton.innerText
    Next HTMLButton
    
    HTMLButtons(0).Click
    
End Sub
Sub BrowseToExchangeRates()
    Dim IE As New SHDocVw.InternetExplorer
    Dim HTMLDoc As MSHTML.HTMLDocument
    Dim HTMLInput As MSHTML.IHTMLElement
    Dim HTMLButtons As MSHTML.IHTMLElementCollection
    Dim HTMLButton As MSHTML.IHTMLElement
    
    IE.Visible = True
    IE.Navigate "x-rates.com"
    
    Do While IE.ReadyState <> READYSTATE_COMPLETE
    Loop
    
    Set HTMLButtons = IE.Document

    For Each HTMLButton In HTMLButtons
        Debug.Print HTMLButton.className, HTMLButton.tagName, HTMLButton.ID, HTMLButton.innerText
    Next HTMLButton
    
    HTMLButtons(0).Click
    
End Sub
'voy en el minuto 30


