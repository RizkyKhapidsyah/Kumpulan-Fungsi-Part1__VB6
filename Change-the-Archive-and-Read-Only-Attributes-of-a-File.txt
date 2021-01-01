Attribute VB_Name = "Module3"
'Clears the vbReadOnly attribute.

 SetAttr strFileName, GetAttr(strFileName) And (Not vbReadOnly)

'Clears the vbArchive attribute.

 SetAttr strFileName, GetAttr(strFileName) And (Not vbArchive)
