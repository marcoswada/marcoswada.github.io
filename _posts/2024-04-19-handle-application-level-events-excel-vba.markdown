---
layout: post
title:  "Handle application level events in Excel VBA"
date:   2024-04-19 13:10:00 +0900
categories: vba excel event-handling
---

# How to handle application level events on Excel VBA

When you need to handle events like creating a new workbook or opening a workbook, you will need to handle application level events.

It is done creating a Class Module that holds the event handling procedures like this:

clsAppEvents.cls
{% highlight vb %}
Public WithEvents App As Application

Private Sub App_NewWorkbook(ByVal Wb As Workbook)
    Debug.Print "Private Sub App_NewWorkbook(ByVal Wb As Workbook): Wb.Name = " & Wb.Name
End Sub
{% endhighlight %}

But in order to work, it need to instantiated. If you want it to be done on opening the file, you can do it in a Sub named auto_open() or in the Workbook_Open event handler. 

{% highlight vb %}
Dim AppEvents As New clsAppEvents
Sub auto_open()
    Set AppEvents.App = Application
End Sub
{% endhighlight %}