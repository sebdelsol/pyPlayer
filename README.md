pyPlayer
========

Python directshow movie player

* plays nearly every formats thanks to LAV filters and XYsubFilter
* fullscreen exclusive mode thanks to madVR
* dvd support (with dvdNavigator filter on windows 7)
* automatic aspect ratio
* favorites
* default languages
* smooth and antialiased OSD
* audio and subtitles tracks selection
* log scale volume
* seamless seek in playlist
* movie rotation
* filter graph is registered in the ROT to be viewed with graphedit
* multi threaded
* embedded in any window

Dependencies
------------

* directshow filters : LAV filters, XYsubFilter, Reclock, madVR
* direct3d9 (needs d3dx9_43.dll in system32)
* python modules: pywin32, comtypes, unidecode, psyco
 
How to use
----------

* register the OSD COM object with osd.py /regserver
* see miniplayer.py

