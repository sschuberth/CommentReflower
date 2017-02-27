# Comment Reflower for Visual Studio

Comment Reflower is an add-in for Microsoft Visual Studio that provides configurable automatic reformatting of block comments, including XML comments. The program was originally created by Ian Nowland for Visual Studio 2003 and 2005. These versions are still available at the [official website](http://commentreflower.sourceforge.net/).

Sadly, the last official update was posted in January 2006 because Ian Nowland has since [changed platforms](http://stackoverflow.com/questions/1837717/comment-reflower-for-visual-studio/3225417#3225417). The last official version of Comment Reflower is not compatible with Visual Studio 2008 or later, apparently due to changes in the extensibility architecture.

Terrified that one might now have to format XML comments manually, [Christoph Nahr](http://www.kynosarges.de/CommentReflower.html) tinkered with the 2005 distribution until he got it to work with Visual Studio 2008 and 2010/2012. Note that this port is not only unofficial but also lacks the MSI setup package of previous versions, although there is a separate binary package for XCopy deployment. Both packages include a [ReadMe](http://www.kynosarges.de/project/CommentReflower/ReadMe.html) file that describes the deployment and operation of the add-in.

**Note:** The Express editions of Visual Studio don’t support any plug-ins, so you need Standard edition or better to run Comment Reflower. And starting with Visual Studio 2015, Microsoft has [deprecated add-ins](https://msdn.microsoft.com/en-us/library/dn246938.aspx) entirely, so Comment Reflower won’t work on any edition.
