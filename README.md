# com4e
Automatically exported from code.google.com/p/com4e !

Introduction

COM4e is a small yet powerfull framework for manipulating java instances from -potentially- COM objects.
History

I wrote COM4e's ancestor some time ago because I had to manipulate java object instances from javascript. Technically, we had an eclipse RCP application containing an SWT Browser displaying HTML files. And these HTML needed to call functions from the RCP container to share data with other applications on that desktop.

Whatever you may think about that architecture, it was the client's request that the communication be done through javascript.

Here's what I did :

    I discovered that the WebBrowser control ("IE's activeX") let's its container provide the implementation for a special javascript object called window.external.
    To do so, COM developpers create a COM object implementing the IDispatch interface, register it using a special method on the WebBrowser control, and manipulate window.external in javascript directly, or from COM native code.
    I decided I could provide some IDIspatch implementation of mine that would act as a transparent proxy to a java POJO.
    The POJO provides the methods and fields.
    COM and IDispatch configuration is provided through annotations.
    The framework handles all the rest automagically. 

COM4e was born. COM4e is the core framework about exposing an annotated POJO through IDispatch. COM4e has no link or dependency with the WebBrowser component. The special SWT glue to create a SWT Browser-like control that exposes the window.external might be found somewhere else (someday) (maybe) :)
Some sample code will make things clear

Some fake COM interface that's easily written as an annotated POJO

public class MyPureJavaCOMInterface extends ComDispatch {
        
        @DispMethod(id=1001,name = "sayHelloTo")
        public String sayHelloTo(String name)
        {
                return String.format("Hello, %s!", name);
        }       
}

and a sample eclipse view

public class TestViewPart extends ViewPart {

        MyCOMAwareBrowser browser;
        
        @Override
        public void createPartControl(Composite parent) {
                browser = new MyCOMAwareBrowser(parent, SWT.NONE);
                
                browser.setWindowExternal(new MyPureJavaCOMInterface());
                
                browser.get("http://some.domain/test.html", null, 5000);
        }

with test.html containing:

<div id="message"></div>
<script>
        var msg = "Sorry, it doesn't work.";
        if(window.external) {
                msg = "Java says '" + window.external.sayHelloTo("World") + "'";
        }
        document.getElementById("message").innerHTML = msg;
</script>
