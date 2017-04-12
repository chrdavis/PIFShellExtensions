Windows Shell extension to allow image types such as Portable Pixmap Format (PPM), Portable Graymap Format (PGM) and Portable Bitmap Format (PBM) to get thumbnail, properties and data object support in Windows Explorer.  Both ascii and binary versions of the above formats are supported.  Also, the generic portable anymap format (PNM) is also covered, which could be any of the previous 3 formats.

Feel free to donate if you find this useful.

[![Donate](https://www.paypalobjects.com/en_US/i/btn/btn_donate_LG.gif)](https://www.paypal.com/cgi-bin/webscr?cmd=_donations&business=727QPFYPDHRKU&lc=US&item_name=Chris%20Davis&currency_code=USD&bn=PP%2dDonationsBF%3abtn_donateCC_LG%2egif%3aNonHosted)

### Download
[Latest 32 or 64 bit versions](https://github.com/chrdavis/PIFShellExtensions/releases/latest)

I originally wrote this Shell extension almost 10 years ago while in graduate school.  I got tired of trying to figure out which PPM image I generated in my project was which because Windows Explorer didn't natively render thumbnails for PPM file types.  This extension will render thumbnails and properties (width, height, dimensions, bit depth) which will display in Windows Explorer.  It also implements a data handler extension which will permit copying the file from Explorer and pasting in applications such as MSPaint or Photoshop.  

Documentation on the image formats can be found at: https://en.wikipedia.org/wiki/Netpbm_format 

![Image description](/images/PIFShellExtensionDemo2.jpg)

What the experience in Windows is without this extension

![Image description](/images/PIFShellExtensionDemo2_before.jpg)
