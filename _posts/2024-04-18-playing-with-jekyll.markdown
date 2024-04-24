---
layout: post
title:  "Playing with Jekyll"
date:   2024-04-18 08:35:33 +0900
categories: test jekyll highlighting
---

Testing support for code snippets:

### Ruby
{% highlight ruby %}
def print_hi(name)
  puts "Hi, #{name}"
end
print_hi('Tom')
#=> prints 'Hi, Tom' to STDOUT.
{% endhighlight %}

### C Language
{% highlight c %}
void hi (char [] name) {
  // prints 'Hi, Tom' to STDOUT.
  printf("Hi, %s", name);
}
{% endhighlight %}

### Visual Basic
{% highlight vb %}
Public Sub hi (name as string)
' prints 'Hi, Tom' to DEBUG Terminal.
  Debug.Print "Hi, " & name
End Sub
{% endhighlight %}

### Python
{% highlight python %}
def hi (name):
  #prints Hi, Tom to console
  print ("Hi, " + name)
{% endhighlight %}

Available languages are available at [Rouge repository][rouge-list].

Check out the [Jekyll docs][jekyll-docs] for more info on how to get the most out of Jekyll. File all bugs/feature requests at [Jekyll’s GitHub repo][jekyll-gh]. If you have questions, you can ask them on [Jekyll Talk][jekyll-talk].

[jekyll-docs]: https://jekyllrb.com/docs/home
[jekyll-gh]:   https://github.com/jekyll/jekyll
[jekyll-talk]: https://talk.jekyllrb.com/
[rouge-list]: https://github.com/rouge-ruby/rouge/wiki/List-of-supported-languages-and-lexers