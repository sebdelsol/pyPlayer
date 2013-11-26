pyPlayer
========

Python directshow movie player

* plays nearly evey format thanks to LAV filters and XYsubFilter
* fullscreen exclusive mode thanks to madVR
* dvd support (windows 7 only)
* automatic aspect ratio
* favorites
* default languages
* smooth and antialiased OSD
* audio and subtitles tracks selection
* log scale volume
* seamless seek in playlist
* movie rotation
* filter graph is viewable with graphedit
* multi threaded
* embedded in any window

Dependencies
------------

* directshow filters : LAV filters, XYsubFilter, Reclock, madVR
* direct3d9 (d3dx9_43.dll)
* python modules: pywin32, comtypes, unidecode, psyco
 
How to use
----------

* register the OSD COM object with osd.py /regserver
* see miniplayer.py

