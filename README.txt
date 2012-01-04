-------------------------------------------


All code (c)2008 Moksa Media all rights reserved
Developer: Andrew Hughes

This file is part of SubtitleTemplateExcel.

SubtitleTemplateExcel is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

SubtitleTemplateExcel is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with SubtitleTemplateExcel.  If not, see <http://www.gnu.org/licenses/>.


-------------------------------------------

SubtitleTemplateExcel is a Microsoft Excel Spreadsheet with a number of Visual Basic macros
added to make working with timecoded subtitle scripts much easier.

Features include:
- use of "." to mean "00"
- validation of proper timecode format
- use and recognition of drop frame and non-drop frame timecodes
- ensuring that subtitle timecodes do not overlap
- automatic calculation of subtitle duration and color coded
  highlighting of short/med/long timecodes (to allow for easy
  identification of abnormally short subtitles)
- ability to export timecodes as STL, XML for Flash, or SUB for Youtube
- ability to shift all/selected timecodes by a fixed amount

Known Issues:
- timecode validation of large spreadsheets is very slow on the mac; I
  may be doing some redundant validation but the PC implementation of the
  Visual Basic interpreter also appears to be much faster

I have included two files called "Sheet Macros - Reference Only" and "Module Macros - Reference Only"
whose purpose is simply to allow github users to view the Visual Basic source embedded in the
Excel spreadsheet without having to download the repository--BUT THESE ARE NOT USED AT ALL
BY THE SPREADSHEET ITSELF.

