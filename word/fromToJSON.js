/*
 * (c) Copyright Ascensio System SIA 2010-2019
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA at 20A-12 Ernesta Birznieka-Upisha
 * street, Riga, Latvia, EU, LV-1050.
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * Pursuant to Section 7(b) of the License you must retain the original Product
 * logo when distributing the program. Pursuant to Section 7(e) we decline to
 * grant you any rights under trademark law for use of our trademarks.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

"use strict";
(function(window, undefined){

	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

	function private_PtToMM(pt)
	{
		return 25.4 / 72.0 * pt;
	}
	function private_Twips2MM(twips)
	{
		return 25.4 / 72.0 / 20 * twips;
	}
	function private_GetDrawingDocument()
	{
		return editor.WordControl.m_oLogicDocument.DrawingDocument;
	}
	function private_EMU2MM(EMU)
	{
		return EMU / 36000.0;
	}
	function private_MM2EMU(MM)
	{
		return MM * 36000.0;
	}
	function private_GetLogicDocument()
	{
		return editor.WordControl.m_oLogicDocument;
	}
	function private_GetStyles()
	{
		var oLogicDocument = private_GetLogicDocument();

		return oLogicDocument instanceof AscCommonWord.CDocument ? oLogicDocument.Get_Styles() : oLogicDocument.globalTableStyles;
	}
	function private_MM2Twips(mm)
	{
		return mm / (25.4 / 72.0 / 20);
	}
	function private_Twips2Px(twips)
	{
		return twips * (4 / 3 / 20);
	}
	function private_Px2Twips(px)
	{
		return px / (4 / 3 / 20);
	}
	/**
	 * Get the first Run in the array specified.
	 * @typeofeditors ["CDE"]
	 * @param {Array} firstPos - first doc pos of element
	 * @param {Array} secondPos - second doc pos of element
	 * @return {1 || 0 || - 1}
	 * If returns 1  -> first element placed before second
	 * If returns 0  -> first element placed like second
	 * If returns -1 -> first element placed after second
	 */
	function private_checkRelativePos(firstPos, secondPos)
	{
		for (var nPos = 0, nLen = Math.min(firstPos.length, secondPos.length); nPos < nLen; ++nPos)
		{
			if (!secondPos[nPos] || !firstPos[nPos] || firstPos[nPos].Class !== secondPos[nPos].Class)
				return 1;

			if (firstPos[nPos].Position < secondPos[nPos].Position)
				return 1;
			else if (firstPos[nPos].Position > secondPos[nPos].Position)
				return -1;
		}

		return 0;
	}
	function private_MM2Pt(mm)
	{
		return mm / (25.4 / 72.0);
	}

	function GetRectAlgnStrType(nAlgnType)
	{
		var sAlgnType = undefined;
		switch (nAlgnType)
		{
			case AscCommon.c_oAscRectAlignType.b:
				sAlgnType = "b";
				break;
			case AscCommon.c_oAscRectAlignType.bl:
				sAlgnType = "bl";
				break;
			case AscCommon.c_oAscRectAlignType.br:
				sAlgnType = "br";
				break;
			case AscCommon.c_oAscRectAlignType.ctr:
				sAlgnType = "ctr";
				break;
			case AscCommon.c_oAscRectAlignType.l:
				sAlgnType = "l";
				break;
			case AscCommon.c_oAscRectAlignType.r:
				sAlgnType = "r";
				break;
			case AscCommon.c_oAscRectAlignType.t:
				sAlgnType = "t";
				break;
			case AscCommon.c_oAscRectAlignType.tl:
				sAlgnType = "tl";
				break;
			case AscCommon.c_oAscRectAlignType.tr:
				sAlgnType = "tr";
				break;
		}

		return sAlgnType;
	}
	function GetRectAlgnNumType(sAlgnType)
	{
		var nAlgnType = undefined;
		switch (sAlgnType)
		{
			case "b":
				nAlgnType = AscCommon.c_oAscRectAlignType.b;
				break;
			case "bl":
				nAlgnType = AscCommon.c_oAscRectAlignType.bl;
				break;
			case "br":
				nAlgnType = AscCommon.c_oAscRectAlignType.br;
				break;
			case "ctr":
				nAlgnType = AscCommon.c_oAscRectAlignType.ctr;
				break;
			case "l":
				nAlgnType = AscCommon.c_oAscRectAlignType.l;
				break;
			case "r":
				nAlgnType = AscCommon.c_oAscRectAlignType.r;
				break;
			case "t":
				nAlgnType = AscCommon.c_oAscRectAlignType.t;
				break;
			case "tl":
				nAlgnType = AscCommon.c_oAscRectAlignType.tl;
				break;
			case "tr":
				nAlgnType = AscCommon.c_oAscRectAlignType.tr;
				break;
		}

		return nAlgnType;
	}

	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// End of private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    function WriterToJSON()
	{
		this.layoutsMap     = {};
		this.mastersMap     = {};
		this.notesMasterMap = {};
		this.themesMap      = {};
	}
    //----------------------------------------------------------export----------------------------------------------------
    window['AscCommon']       = window['AscCommon'] || {};
    window['AscFormat']       = window['AscFormat'] || {};
	window['AscCommon'].WriterToJSON   = WriterToJSON;
	window['AscCommon'].ReaderFromJSON = ReaderFromJSON;
})(window);



