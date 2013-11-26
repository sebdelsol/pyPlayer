pyPlayer
========

Python directshow movie player

* plays nearly every formats with LAV filters
* fullscreen exclusive mode with madVR
* dvd support with dvdNavigator filter
* automatic aspect ratio
* bookmarks
* default languages
* responsive antialiased OSD
* audio and subtitles tracks selection
* log scale volume
* seamless seek in playlist
* movie rotation
* filter graph is registered in the ROT to be viewed with graphedit
* multi threaded
* could be embedded in any window

Dependencies
------------

* directshow filters : LAV filters, XYsubFilter, Reclock, madVR
* direct3d9 (d3dx9_43.dll)
* python modules: pywin32, comtypes, unidecode, psyco
 
How to use
----------

* register the OSD COM interface with osd.py /regserver
* see miniplayer.py

Limitations
-----------
* no windowed mode (move and resize events are not handled)
* no fallback when the graph can't be build
* no mouse handling
* no filter choice (because those chosen are the best at the moment IMHO)
* iso639-2 languages decoded in french only
