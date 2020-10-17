#### VBCorLib.res File Contents ####

VBCorLib contains three custom resources that contain culture and encoding information.

* VBCultures.nlp
* EncodingInfo.nlp
* CodePageLookup.bin

##### VBCultures.nlp
Contains all the culture information used to perform date/time and numeric formatting.

Adding the file as a Visual Basic resource through the Resource Editor:
1. Add file as a custom resource.
2. Rename resource to "CULTUREINFO".
3. Set resource number to 101.

##### EncodingInfo.nlp
Contains information used to encode and decode strings in various encoding schemes.

Adding the file as a Visual Basic resource through the Resource Editor:
1. Add file as a custom resource.
2. Rename resource to "ENCODINGINFO".
3. Set resource number to 101.

##### CodePageLookup.bin
Contains various names encodings may be referenced other than the Encoding.WebName to
help improve accessing the many encodings without needing to know the exact name.

Adding the file as a Visual Basic resource through the Resource Editor:
1. Add file as a custom resource.
2. Rename resource to "CODEPAGELOOKUP".
3. Set resource number to 101.
