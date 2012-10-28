** VBCorLib 2.3 **
- Updated ASM code to not trigger DEP protection.
- Updated ArrayList to include the ArrayList.Adapter method, wrapping an IList object.
- Fixed the StringBuilder.Remove method.
- Updated Console class to automatically disable the close button on a command prompt when working in the IDE.
- Added TryParse to TimeSpan.
- Added TryParse to cDateTime.


** VBCorLib 2.2 **
- Added the BigInteger class
- Fixed Queue class. Incorrectly zeroed out memory, causing a crash.
- Fixed Path class. Incorrectly handled files begining with a period.
- Fixed Path class. Incorrectly handled directory names that include '..' to move up a directory.
- Fixed HebrewCalendard class. Incorrectly handled day 355 of a leap year.
- Fixed JapaneseCalender class. Incorrectly handled the start of each year after the first of an era.
- Fixed DSAParameter class. Now includes the J parameter on XML export.


** New in Version 2.1 **
- Added Wide Character API support to INIFile and INISectionWriter classes
- Added EmptyGuid to the GuidStatic class.
- Fixed the Wide Character support for file API calls used on NT machines.
- Fixed the return types for SHA384Static and SHA512Static create methods.



** New in Version 2.0 **
Many many cryptography classes have been added.
This list is only a portion of all the additions and changes to the library.

Symmetric Algorithms
--------------------
Rijndael
DES
Triple DES
RC2

Asymmetric Algorithms
---------------------
RSA

Asymmetric signing and verification algorithms
----------------------------------------------
RSA
DSA

Hash Algorithms
---------------
SHA-1
SHA-256
SHA-384
SHA-512
MD5
RIPEMD160

Message Authentication Code Algorithms
--------------------------------------
HMACSHA-1
HMACSHA-256
HMACSHA-384
HMACSHA-512
HMACMD5
HMACRIPEMD160
TripleDES Mac

Password Generators
-------------------
Rfc2898DeriveBytes

Utility Classes
---------------
CryptoStream
CryptoConfig
SecurityElement
RNGCryptoServiceProvider


Unicode API support
-------------------
All of the file, directory and registry classes support Unicode directly on NT
machines. This allows for long file names (over 32k).

VBCultures.nlp and Resource File
--------------------------------
The VBCultures.nlp file has been included as a resource to be compiled directly
into the library. This is to reduce the dependency of handling multiple files
for distribution.

There are also several code updates through the library.


If you have any questions, please visit the VBCorLib website:
http://www.kellyethridge.com/vbcorlib
