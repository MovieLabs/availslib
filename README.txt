README - availslib
------------------

0) Introduction: "availslib" is a Java library with methods for
   converting EMA avails data contained within an excel spreadsheet
   into an XML representation.  Both "SAX" (i.e. row at a time) and
   DOM (i.e. worksheet at a time) styles of conversion are supported.
   In addition to converting files, extensive validation is performed
   to detect common errors in the Spreadsheet data.  The XML documents
   created will validate against the Avails-related schema.

1) Source control: these files are maintained in a github repository.
   You can use the following shell command (either 'nix or Win32) to
   download the files therefrom:

> mkdir <your target directory> (e.g. availslib)
> cd <your target directory>
> git init
> git remote add github https://github.com/pgj-ml/availslib.git
> git fetch github

3) Building availslib: Eclipse was used for initial development;
   however any IDE should work equally well.  In addition to the
   source files you will need the following libraries:
   
   - JRE System Library (JavaSE-1.8 recommended)
   - Apache Poi (https://poi.apache.org/); used the following jars:
       * poi-3.13-20150929.jar
       * poi-ooxml-3.13-20150929.jar
       * poi-ooxml-schemas-3.13-20150929.jar
   - Apache XML Beans (https://xmlbeans.apache.org/); used the following jar:
       * xmlbeans-2.6.0.jar
   - Apache Apache Log4j 2 (http://logging.apache.org/log4j/2.x/); used the
     following jars:
       * apache-log4j-2.4.1-bin/log4j-api-2.4.1.jar
       * apache-log4j-2.4.1-bin/log4j-core-2.4.1.jar
     
4) Documentation: the source code has JavaDoc annotations.  You need
   to install a Javadoc processor to generate html-based
   documentation.  This is usually provided by your IDE, but you can
   refer here for more information:

   http://www.oracle.com/technetwork/articles/java/index-jsp-135444.html

5) Test and verification: to perform a quick test of avilslib, use the
   "availstool" command-line utility.  This can also be found on
   github at:

   https://github.com/pgj-ml/availslib.git

6) for further information: contact Paul Jensen (pgj@movielabs.com)
