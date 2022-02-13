#excel #vba #designpatterns #observerpattern

As a programmer, have you ever given a thought to how chat applications are being programmed? How Worksheet_Change event gets fired whenever someone changes cell value? How do all the subscribers or receivers get notified instantly whenever the sender or publisher post something?

Some of you know that I & Ismail has started exploring design patterns within VBA and it started well and has revealed many secrets of some of the best app designs. Earlier I have made a post on the Strategy pattern and Ismail has made a post on the command pattern. You can find here the link to both posts. 

https://www.linkedin.com/posts/md-ismail-hosen-b77500135_github-1504168command-pattern-vba-activity-6888136703740145664-9UK8/
https://www.linkedin.com/posts/kamalbharakhda_travel-planner-using-strategy-patternxlsm-activity-6888122845180895232-ozGE/

So, design patterns are kind of proven strategies to design specific and operation-oriented applications and services. It gives you a ready framework that you can implement easily. 

The observer pattern is a kind of pattern that deals with specific operations which is nothing but the mechanism of the notifier. it has two interface class, one is the subscriber/observer means who is watching or observing something and it should be notified if something change in whatever they are observing while another interface is of Subject or Publisher or Sender who is getting watched by the subscribers. 

In the observer pattern, each of the observers has to register first with the subject or publisher class so the publisher will know whom to notify whenever something changes. 
I have designed a simple chat application to demonstrate key design points of the observer pattern. Kindly look at the attached video as well so you will get how is it easier to build something like that. 
