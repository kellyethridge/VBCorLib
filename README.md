# VBCorLib

## Overview

VBCorLib is a Visual Basic 6 implementation of many classes found in the .NET framework. The classes
within VBCorLib can be used nearly identically as the .NET counterpart. This allows for easy data
sharing between a .NET application and VB6.

* Provides several collection types: ArrayList, Stack, Queue and Hashtable.
* Provides several encryption algorithms: Rijndael, RSA, TripleDES, DES.
* Provides many hashing algorithms: SHA1, SHA256, SHA384, SHA516, RIPMED160, MD5.
* Sign and verify data using HMAC.
* Provides easy access to many encodings for text and file handling: UTF8, UTF7, UTF16, UTF32, and Windows supported encodings.
* Easy String, Array and Date manipulation with a variety of classes.
* Manipulate files with a variety of file handling classes.
* Handles files larger than 2 gigs.
* Provides a BigInteger to perform large calculations.
* Utilize signed 64-bit integers.
* Provides easy access to a console window.
* And much more...

## Documentation (Updated 2020-10-11)

* Browse online at <http://www.kellyethridge.com/vbcorlib/doc/CorLib.html>.
* Download offline from <http://www.kellyethridge.com/vbcorlib/doc/CorLib.chm>.

## Blog

There is a blog that I attempt to update on occasion at <http://vbcorlib.blogspot.com/>.

## Maintaining

**WARNING:** In order to get the correct windows EOLs that the VB6 IDE demands for repository users
that access the sources using [GitHub's Subversion checkout support](
https://help.github.com/en/articles/support-for-subversion-clients), all maintainers (including PR
requestors) must set their git autocrlf configuration to false:

```shell
git config --global core.autocrlf false
```
