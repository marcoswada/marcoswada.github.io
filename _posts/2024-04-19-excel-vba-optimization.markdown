---
layout: post
title:  "VBA Optimization in Excel"
date:   2024-04-19 13:30:00 +0900
categories: vba excel optimization
---


# VBA optimization in Excel

There is nothing more frustrating than your new code strugling to do some task that wasn't even supposed to be so complex.

You can use some trick like snippet below:

{% highlight vb %}
Sub OptimizeVBA(isOn As Boolean)
    Application.ScreenUpdating = Not isOn
    Application.Calculation = IIf(isOn, xlCalculationManual, xlCalculationAutomatic)
    Application.EnableEvents = Not isOn
    ActiveSheet.displaypagebreak = Not isOn
End Sub
{% endhighlight %}

It can help you sometimes but when you manipulate a lot of data in excel sheets, there is another problem: *abstraction*. It was meant to be something good, but, there is a lot hidden in a simple operation of assigning a value into a cell, so if you execute the snippet below, that fills 1 million of cells with random numbers (which in any programming language would take a fraction of a second, it could take several seconds, depending on you hardware configuration).

{% highlight vb %}
Sub randomNumbers()
    Randomize Timer
    Dim i As Integer
    Dim j As Integer
    Dim t As Double
    t = Timer
    For i = 1 To 1000
        For j = 1 To 1000
            Sheets(1).Cells(i, j) = Int(Rnd * 1000) + 1
        Next
    Next
    Debug.Print Timer - t & " seconds elapsed."
End Sub
{% endhighlight %}

Well, the code above took me *103.25* seconds on a 4th gen Intel i3 and *36.38* seconds on a 6th gen i7.

Let's take another approach. You can do the same task assigning the values into a memory array and then assign the array itself into a range of cells.

{% highlight vb %}
Sub optimizedRandomNumbers()
    Randomize Timer
    Dim i As Integer
    Dim j As Integer
    Dim arr(1 To 1000, 1 To 1000) As Integer
    Dim t As Double
    t = Timer
    For i = 1 To 1000
        For j = 1 To 1000
            arr(i, j) = Int(Rnd * 1000) + 1
        Next
    Next
    Sheets(1).Cells(1, 1).Resize(1000, 1000) = arr
    Debug.Print Timer - t & " seconds elapsed."
End Sub
{% endhighlight %}

Wow! The code above took me *2.75* seconds on the 4th gen i3 and *1.04* second on the 6th gen i7. (And yes, I know that we usually work with 0 based arrays)

This huge difference is because Excel abstracts so much stuff into the simple operation of assigning a value into a cell like cell formats, data typing, etc. that assigning some data into a cell individually, it takes all those processing that can be done at once when you assign an entire array into a range.