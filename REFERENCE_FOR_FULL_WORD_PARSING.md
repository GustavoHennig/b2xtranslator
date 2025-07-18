Here you will find the main part of a fully functional Word document parser implemented in C, this handles all scenarios of fastsaved documents.



## clx.c
```c
/* wvWare
 * Copyright (C) Caolan McNamara, Dom Lachowicz, and others
 *
 * This program is free software; you can redistribute it and/or
 * modify it under the terms of the GNU General Public License
 * as published by the Free Software Foundation; either version 2
 * of the License, or (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program; if not, write to the Free Software
 * Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA
 * 02111-1307, USA.
 */

#ifdef HAVE_CONFIG_H
#include "config.h"
#endif

#include <stdlib.h>
#include <stdio.h>
#include <string.h>
#include "wv.h"

void
wvReleaseCLX (CLX * clx)
{
    U16 i;
    for (i = 0; i < clx->grpprl_count; i++)
	wvFree (clx->grpprl[i]);
    wvFree (clx->grpprl);
    wvFree (clx->cbGrpprl);
    wvReleasePCD_PLCF (clx->pcd, clx->pos);
}

void
wvBuildCLXForSimple6 (CLX * clx, FIB * fib)
{
    wvInitCLX (clx);
    clx->nopcd = 1;;

    clx->pcd = (PCD *) wvMalloc (clx->nopcd * sizeof (PCD));
    clx->pos = (U32 *) wvMalloc ((clx->nopcd + 1) * sizeof (U32));

    clx->pos[0] = 0;
    clx->pos[1] = fib->ccpText;

    wvInitPCD (&(clx->pcd[0]));
    clx->pcd[0].fc = fib->fcMin;

    /* reverse the special encoding thing they do for word97 
       if we are using the usual 8 bit chars */

    if (fib->fExtChar == 0)
      {
	  clx->pcd[0].fc *= 2;
	  clx->pcd[0].fc |= 0x40000000UL;
      }

    clx->pcd[0].prm.fComplex = 0;
    clx->pcd[0].prm.para.var1.isprm = 0;
    /*
       these set the ones that *I* use correctly, but may break for other wv
       users, though i doubt it, im just marking a possible firepoint for the
       future
     */
}

/*
The complex part of a file (CLX) is composed of a number of variable-sized
blocks of data. Recorded first are any grpprls that may be referenced by the
plcfpcd (if the plcfpcd has no grpprl references, no grpprls will be
recorded) followed by the plcfpcd. Each block in the complex part is
prefaced by a clxt (clx type), which is a 1-byte code, either 1 (meaning the
block contains a grpprl) or 2 (meaning this is the plcfpcd). A clxtGrpprl
(1) is followed by a 2-byte cb which is the count of bytes of the grpprl. A
clxtPlcfpcd (2) is followed by a 4-byte lcb which is the count of bytes of
the piece table. A full saved file will have no clxtGrpprl's.
*/
void
wvGetCLX (wvVersion ver, CLX * clx, U32 offset, U32 len, U8 fExtChar,
	  wvStream * fd)
{
    U8 clxt;
    U16 cb;
    U32 lcb, i, j = 0;

    wvTrace (("offset %x len %d\n", offset, len));
    wvStream_goto (fd, offset);

    wvInitCLX (clx);

    while (j < len)
      {
	  clxt = read_8ubit (fd);
	  j++;
	  if (clxt == 1)
	    {
		cb = read_16ubit (fd);
		j += 2;
		clx->grpprl_count++;
		clx->cbGrpprl =
		    (U16 *) realloc (clx->cbGrpprl,
				     sizeof (U16) * clx->grpprl_count);
		clx->cbGrpprl[clx->grpprl_count - 1] = cb;
		clx->grpprl =
		    (U8 **) realloc (clx->grpprl,
				     sizeof (U8 *) * (clx->grpprl_count));
		clx->grpprl[clx->grpprl_count - 1] = (U8 *) wvMalloc (cb);
		for (i = 0; i < cb; i++)
		    clx->grpprl[clx->grpprl_count - 1][i] = read_8ubit (fd);
		j += i;
	    }
	  else if (clxt == 2)
	    {
		if (ver == WORD8)
		  {
		      lcb = read_32ubit (fd);
		      j += 4;
		  }
		else
		  {
		      wvTrace (("Here so far\n"));
#if 0
		      lcb = read_16ubit (fd);	/* word 6 only has two bytes here */
		      j += 2;
#endif

		      lcb = read_32ubit (fd);	/* word 6 specs appeared to have lied ! */
		      j += 4;
		  }
		wvGetPCD_PLCF (&clx->pcd, &clx->pos, &clx->nopcd,
			       wvStream_tell (fd), lcb, fd);
		j += lcb;

		if (ver <= WORD7)	/* MV 28.8.2000 Appears to be valid */
		  {
#if 0
		      /* DANGER !!, this is a completely mad attempt to differenciate 
		         between word 95 files that use 16 and 8 bit characters. It may
		         not work, it attempt to err on the side of 8 bit characters.
		       */
		      if (!(wvGuess16bit (clx->pcd, clx->pos, clx->nopcd)))
#else
		      /* I think that this is the correct reason for this behaviour */
		      if (fExtChar == 0)
#endif
			  for (i = 0; i < clx->nopcd; i++)
			    {
				clx->pcd[i].fc *= 2;
				clx->pcd[i].fc |= 0x40000000UL;
			    }
		  }
	    }
	  else
	    {
		wvError (("clxt is not 1 or 2, it is %d\n", clxt));
		return;
	    }
      }
}


void
wvInitCLX (CLX * item)
{
    item->pcd = NULL;
    item->pos = NULL;
    item->nopcd = 0;

    item->grpprl_count = 0;
    item->cbGrpprl = NULL;
    item->grpprl = NULL;
}


int
wvGetPieceBoundsFC (U32 * begin, U32 * end, CLX * clx, U32 piececount)
{
    int type;
    if ((piececount + 1) > clx->nopcd)
      {
	  wvTrace (
		   ("piececount is > nopcd, i.e.%d > %d\n", piececount + 1,
		    clx->nopcd));
	  return (-1);
      }
    *begin = wvNormFC (clx->pcd[piececount].fc, &type);

    if (type)
	*end = *begin + (clx->pos[piececount + 1] - clx->pos[piececount]);
    else
	*end = *begin + ((clx->pos[piececount + 1] - clx->pos[piececount]) * 2);

    return (type);
}

int
wvGetPieceBoundsCP (U32 * begin, U32 * end, CLX * clx, U32 piececount)
{
    if ((piececount + 1) > clx->nopcd)
	return (-1);
    *begin = clx->pos[piececount];
    *end = clx->pos[piececount + 1];
    return (0);
}


char *
wvAutoCharset (wvParseStruct * ps)
{
    U16 i = 0;
    int flag;
    char *ret;
    ret = "iso-8859-15";

    /* 
       If any of the pieces use unicode then we have to assume the
       worst and use utf-8
     */
    while (i < ps->clx.nopcd)
      {
	  wvNormFC (ps->clx.pcd[i].fc, &flag);
	  if (flag == 0)
	    {
		ret = "UTF-8";
		break;
	    }
	  i++;
      }

    /* 
       Also if the document fib is not codepage 1252 we also have to 
       assume the worst 
     */
    if (strcmp (ret, "UTF-8"))
      {
	  if (
	      (ps->fib.lid != 0x407) &&
	      (ps->fib.lid != 0x807) &&
	      (ps->fib.lid != 0x409) &&
	      (ps->fib.lid != 0x807) && (ps->fib.lid != 0xC09))
	      ret = "UTF-8";
      }
    return (ret);
}



int
wvQuerySamePiece (U32 fcTest, CLX * clx, U32 piece)
{
    /*
       wvTrace(("Same Piece, %x %x %x\n",fcTest,wvNormFC(clx->pcd[piece].fc,NULL),wvNormFC(clx->pcd[piece+1].fc,NULL)));
       if ( (fcTest >= wvNormFC(clx->pcd[piece].fc,NULL)) && (fcTest < wvNormFC(clx->pcd[piece+1].fc,NULL)) )
     */
    wvTrace (
	     ("Same Piece, %x %x %x\n", fcTest, clx->pcd[piece].fc,
	      wvGetEndFCPiece (piece, clx)));
    if ((fcTest >= wvNormFC (clx->pcd[piece].fc, NULL))
	&& (fcTest < wvGetEndFCPiece (piece, clx)))
	return (1);
    return (0);
}


U32
wvGetPieceFromCP (U32 currentcp, CLX * clx)
{
    U32 i = 0;
    while (i < clx->nopcd)
      {
	  wvTrace (
		   ("i %d: currentcp is %d, clx->pos[i] is %d, clx->pos[i+1] is %d\n",
		    i, currentcp, clx->pos[i], clx->pos[i + 1]));
	  if ((currentcp >= clx->pos[i]) && (currentcp < clx->pos[i + 1]))
	      return (i);
	  i++;
      }
    wvTrace (("cp was not in any piece ! \n", currentcp));
    return (0xffffffffL);
}

U32
wvGetEndFCPiece (U32 piece, CLX * clx)
{
    int flag;
    U32 fc;
    U32 offset = clx->pos[piece + 1] - clx->pos[piece];

    wvTrace (("offset is %x, befc is %x\n", offset, clx->pcd[piece].fc));
    fc = wvNormFC (clx->pcd[piece].fc, &flag);
    wvTrace (("fc is %x, flag %d\n", fc, flag));
    if (flag)
	fc += offset;
    else
	fc += offset * 2;
    wvTrace (("fc is finally %x\n", fc));
    return (fc);
}

/*
1) search for the piece containing the character in the piece table.

2) Then calculate the FC in the file that stores the character from the piece
    table information.
*/
U32
wvConvertCPToFC (U32 currentcp, CLX * clx)
{
    U32 currentfc = 0xffffffffL;
    U32 i = 0;
    int flag;

    while (i < clx->nopcd)
      {
	  if ((currentcp >= clx->pos[i]) && (currentcp < clx->pos[i + 1]))
	    {
		currentfc = wvNormFC (clx->pcd[i].fc, &flag);
		if (flag)
		    currentfc += (currentcp - clx->pos[i]);
		else
		    currentfc += ((currentcp - clx->pos[i]) * 2);
		break;
	    }
	  i++;
      }

    if (currentfc == 0xffffffffL)
      {
	  i--;
	  currentfc = wvNormFC (clx->pcd[i].fc, &flag);
	  if (flag)
	      currentfc += (currentcp - clx->pos[i]);
	  else
	      currentfc += ((currentcp - clx->pos[i]) * 2);
	  wvTrace (("flaky cp to fc conversion underway\n"));
      }

    return (currentfc);
}

struct test {
    U32 fc;
    U32 offset;
};

int
compar (const void *a, const void *b)
{
    struct test *one, *two;
    one = (struct test *) a;
    two = (struct test *) b;

    if (one->fc < two->fc)
	return (-1);
    else if (one->fc == two->fc)
	return (0);
    return (1);
}

/* 
In word 95 files there is no flag attached to each
offset as there is in word 97 to tell you that we are
talking about 16 bit chars, so I attempt here to make
an educated guess based on overlapping offsets to
figure it out, If I had some actual information as
the how word 95 actually stores it it would help.
*/

int
wvGuess16bit (PCD * pcd, U32 * pos, U32 nopcd)
{
    struct test *fcs;
    U32 i;
    int ret = 1;
    fcs = (struct test *) wvMalloc (sizeof (struct test) * nopcd);
    for (i = 0; i < nopcd; i++)
      {
	  fcs[i].fc = pcd[i].fc;
	  fcs[i].offset = (pos[i + 1] - pos[i]) * 2;
      }

    qsort (fcs, nopcd, sizeof (struct test), compar);

    for (i = 0; i < nopcd - 1; i++)
      {
	  if (fcs[i].fc + fcs[i].offset > fcs[i + 1].fc)
	    {
		wvTrace (("overlap, my guess is 8 bit\n"));
		ret = 0;
		break;
	    }
      }

    wvFree (fcs);
    return (ret);
}

```

## fib.c
```c
/* wvWare
 * Copyright (C) Caolan McNamara, Dom Lachowicz, and others
 *
 * This program is free software; you can redistribute it and/or
 * modify it under the terms of the GNU General Public License
 * as published by the Free Software Foundation; either version 2
 * of the License, or (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program; if not, write to the Free Software
 * Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA
 * 02111-1307, USA.
 */

#ifdef HAVE_CONFIG_H
#include "config.h"
#endif

#include <stdlib.h>
#include <stdio.h>
#include <string.h>
#include "wv.h"

void
wvInitFIB (FIB * item)
{
    item->wIdent = 0;
    item->nFib = 0;
    item->nProduct = 0;
    item->lid = 0;
    item->pnNext = 0;
    item->fDot = 0;
    item->fGlsy = 0;
    item->fComplex = 0;
    item->fHasPic = 0;
    item->cQuickSaves = 0;
    item->fEncrypted = 0;
    item->fWhichTblStm = 0;
    item->fReadOnlyRecommended = 0;
    item->fWriteReservation = 0;
    item->fExtChar = 0;
    item->fLoadOverride = 0;
    item->fFarEast = 0;
    item->fCrypto = 0;
    item->nFibBack = 0;
    item->lKey = 0;
    item->envr = 0;
    item->fMac = 0;
    item->fEmptySpecial = 0;
    item->fLoadOverridePage = 0;
    item->fFutureSavedUndo = 0;
    item->fWord97Saved = 0;
    item->fSpare0 = 0;
    item->chse = 0;
    item->chsTables = 0;
    item->fcMin = 0;
    item->fcMac = 0;
    item->csw = 0;
    item->wMagicCreated = 0;
    item->wMagicRevised = 0;
    item->wMagicCreatedPrivate = 0;
    item->wMagicRevisedPrivate = 0;
    item->pnFbpChpFirst_W6 = 0;
    item->pnChpFirst_W6 = 0;
    item->cpnBteChp_W6 = 0;
    item->pnFbpPapFirst_W6 = 0;
    item->pnPapFirst_W6 = 0;
    item->cpnBtePap_W6 = 0;
    item->pnFbpLvcFirst_W6 = 0;
    item->pnLvcFirst_W6 = 0;
    item->cpnBteLvc_W6 = 0;
    item->lidFE = 0;
    item->clw = 0;
    item->cbMac = 0;
    item->lProductCreated = 0;
    item->lProductRevised = 0;
    item->ccpText = 0;
    item->ccpFtn = 0;
    item->ccpHdr = 0;
    item->ccpMcr = 0;
    item->ccpAtn = 0;
    item->ccpEdn = 0;
    item->ccpTxbx = 0;
    item->ccpHdrTxbx = 0;
    item->pnFbpChpFirst = 0;
    item->pnChpFirst = 0;
    item->cpnBteChp = 0;
    item->pnFbpPapFirst = 0;
    item->pnPapFirst = 0;
    item->cpnBtePap = 0;
    item->pnFbpLvcFirst = 0;
    item->pnLvcFirst = 0;
    item->cpnBteLvc = 0;
    item->fcIslandFirst = 0;
    item->fcIslandLim = 0;
    item->cfclcb = 0;
    item->fcStshfOrig = 0;
    item->lcbStshfOrig = 0;
    item->fcStshf = 0;
    item->lcbStshf = 0;
    item->fcPlcffndRef = 0;
    item->lcbPlcffndRef = 0;
    item->fcPlcffndTxt = 0;
    item->lcbPlcffndTxt = 0;
    item->fcPlcfandRef = 0;
    item->lcbPlcfandRef = 0;
    item->fcPlcfandTxt = 0;
    item->lcbPlcfandTxt = 0;
    item->fcPlcfsed = 0;
    item->lcbPlcfsed = 0;
    item->fcPlcpad = 0;
    item->lcbPlcpad = 0;
    item->fcPlcfphe = 0;
    item->lcbPlcfphe = 0;
    item->fcSttbfglsy = 0;
    item->lcbSttbfglsy = 0;
    item->fcPlcfglsy = 0;
    item->lcbPlcfglsy = 0;
    item->fcPlcfhdd = 0;
    item->lcbPlcfhdd = 0;
    item->fcPlcfbteChpx = 0;
    item->lcbPlcfbteChpx = 0;
    item->fcPlcfbtePapx = 0;
    item->lcbPlcfbtePapx = 0;
    item->fcPlcfsea = 0;
    item->lcbPlcfsea = 0;
    item->fcSttbfffn = 0;
    item->lcbSttbfffn = 0;
    item->fcPlcffldMom = 0;
    item->lcbPlcffldMom = 0;
    item->fcPlcffldHdr = 0;
    item->lcbPlcffldHdr = 0;
    item->fcPlcffldFtn = 0;
    item->lcbPlcffldFtn = 0;
    item->fcPlcffldAtn = 0;
    item->lcbPlcffldAtn = 0;
    item->fcPlcffldMcr = 0;
    item->lcbPlcffldMcr = 0;
    item->fcSttbfbkmk = 0;
    item->lcbSttbfbkmk = 0;
    item->fcPlcfbkf = 0;
    item->lcbPlcfbkf = 0;
    item->fcPlcfbkl = 0;
    item->lcbPlcfbkl = 0;
    item->fcCmds = 0;
    item->lcbCmds = 0;
    item->fcPlcmcr = 0;
    item->lcbPlcmcr = 0;
    item->fcSttbfmcr = 0;
    item->lcbSttbfmcr = 0;
    item->fcPrDrvr = 0;
    item->lcbPrDrvr = 0;
    item->fcPrEnvPort = 0;
    item->lcbPrEnvPort = 0;
    item->fcPrEnvLand = 0;
    item->lcbPrEnvLand = 0;
    item->fcWss = 0;
    item->lcbWss = 0;
    item->fcDop = 0;
    item->lcbDop = 0;
    item->fcSttbfAssoc = 0;
    item->lcbSttbfAssoc = 0;
    item->fcClx = 0;
    item->lcbClx = 0;
    item->fcPlcfpgdFtn = 0;
    item->lcbPlcfpgdFtn = 0;
    item->fcAutosaveSource = 0;
    item->lcbAutosaveSource = 0;
    item->fcGrpXstAtnOwners = 0;
    item->lcbGrpXstAtnOwners = 0;
    item->fcSttbfAtnbkmk = 0;
    item->lcbSttbfAtnbkmk = 0;
    item->fcPlcdoaMom = 0;
    item->lcbPlcdoaMom = 0;
    item->fcPlcdoaHdr = 0;
    item->lcbPlcdoaHdr = 0;
    item->fcPlcspaMom = 0;
    item->lcbPlcspaMom = 0;
    item->fcPlcspaHdr = 0;
    item->lcbPlcspaHdr = 0;
    item->fcPlcfAtnbkf = 0;
    item->lcbPlcfAtnbkf = 0;
    item->fcPlcfAtnbkl = 0;
    item->lcbPlcfAtnbkl = 0;
    item->fcPms = 0;
    item->lcbPms = 0;
    item->fcFormFldSttbs = 0;
    item->lcbFormFldSttbs = 0;
    item->fcPlcfendRef = 0;
    item->lcbPlcfendRef = 0;
    item->fcPlcfendTxt = 0;
    item->lcbPlcfendTxt = 0;
    item->fcPlcffldEdn = 0;
    item->lcbPlcffldEdn = 0;
    item->fcPlcfpgdEdn = 0;
    item->lcbPlcfpgdEdn = 0;
    item->fcDggInfo = 0;
    item->lcbDggInfo = 0;
    item->fcSttbfRMark = 0;
    item->lcbSttbfRMark = 0;
    item->fcSttbCaption = 0;
    item->lcbSttbCaption = 0;
    item->fcSttbAutoCaption = 0;
    item->lcbSttbAutoCaption = 0;
    item->fcPlcfwkb = 0;
    item->lcbPlcfwkb = 0;
    item->fcPlcfspl = 0;
    item->lcbPlcfspl = 0;
    item->fcPlcftxbxTxt = 0;
    item->lcbPlcftxbxTxt = 0;
    item->fcPlcffldTxbx = 0;
    item->lcbPlcffldTxbx = 0;
    item->fcPlcfhdrtxbxTxt = 0;
    item->lcbPlcfhdrtxbxTxt = 0;
    item->fcPlcffldHdrTxbx = 0;
    item->lcbPlcffldHdrTxbx = 0;
    item->fcStwUser = 0;
    item->lcbStwUser = 0;
    item->fcSttbttmbd = 0;
    item->cbSttbttmbd = 0;
    item->fcUnused = 0;
    item->lcbUnused = 0;
    item->fcPgdMother = 0;
    item->lcbPgdMother = 0;
    item->fcBkdMother = 0;
    item->lcbBkdMother = 0;
    item->fcPgdFtn = 0;
    item->lcbPgdFtn = 0;
    item->fcBkdFtn = 0;
    item->lcbBkdFtn = 0;
    item->fcPgdEdn = 0;
    item->lcbPgdEdn = 0;
    item->fcBkdEdn = 0;
    item->lcbBkdEdn = 0;
    item->fcSttbfIntlFld = 0;
    item->lcbSttbfIntlFld = 0;
    item->fcRouteSlip = 0;
    item->lcbRouteSlip = 0;
    item->fcSttbSavedBy = 0;
    item->lcbSttbSavedBy = 0;
    item->fcSttbFnm = 0;
    item->lcbSttbFnm = 0;
    item->fcPlcfLst = 0;
    item->lcbPlcfLst = 0;
    item->fcPlfLfo = 0;
    item->lcbPlfLfo = 0;
    item->fcPlcftxbxBkd = 0;
    item->lcbPlcftxbxBkd = 0;
    item->fcPlcftxbxHdrBkd = 0;
    item->lcbPlcftxbxHdrBkd = 0;
    item->fcDocUndo = 0;
    item->lcbDocUndo = 0;
    item->fcRgbuse = 0;
    item->lcbRgbuse = 0;
    item->fcUsp = 0;
    item->lcbUsp = 0;
    item->fcUskf = 0;
    item->lcbUskf = 0;
    item->fcPlcupcRgbuse = 0;
    item->lcbPlcupcRgbuse = 0;
    item->fcPlcupcUsp = 0;
    item->lcbPlcupcUsp = 0;
    item->fcSttbGlsyStyle = 0;
    item->lcbSttbGlsyStyle = 0;
    item->fcPlgosl = 0;
    item->lcbPlgosl = 0;
    item->fcPlcocx = 0;
    item->lcbPlcocx = 0;
    item->fcPlcfbteLvc = 0;
    item->lcbPlcfbteLvc = 0;
    wvInitFILETIME (&item->ftModified);
    item->fcPlcflvc = 0;
    item->lcbPlcflvc = 0;
    item->fcPlcasumy = 0;
    item->lcbPlcasumy = 0;
    item->fcPlcfgram = 0;
    item->lcbPlcfgram = 0;
    item->fcSttbListNames = 0;
    item->lcbSttbListNames = 0;
    item->fcSttbfUssr = 0;
    item->lcbSttbfUssr = 0;

    /* Word 2 */
    item->Spare = 0;
    item->rgwSpare0[0] = 0;
    item->rgwSpare0[1] = 0;
    item->rgwSpare0[2] = 0;
    item->fcSpare0 = 0;
    item->fcSpare1 = 0;
    item->fcSpare2 = 0;
    item->fcSpare3 = 0;
    item->ccpSpare0 = 0;
    item->ccpSpare1 = 0;
    item->ccpSpare2 = 0;
    item->ccpSpare3 = 0;
    item->fcPlcfpgd = 0;
    item->cbPlcfpgd = 0;

    item->fcSpare5 = 0;
    item->cbSpare5 = 0;
    item->fcSpare6 = 0;
    item->cbSpare6 = 0;
    item->wSpare4 = 0;
}

void
wvGetFIB (FIB * item, wvStream * fd)
{
    U16 temp16;
    U8 temp8;

    item->fEncrypted = 0;

    wvStream_goto (fd, 0);
#ifdef PURIFY
    wvInitFIB (item);
#endif
    item->wIdent = read_16ubit (fd);
    item->nFib = read_16ubit (fd);

    if ((wvQuerySupported (item, NULL) == WORD2))
      {
	  wvInitFIB (item);
	  wvStream_offset (fd, -4);
	  wvGetFIB2 (item, fd);
	  return;
      }

    if ((wvQuerySupported (item, NULL) == WORD5)
	|| (wvQuerySupported (item, NULL) == WORD6)
	|| (wvQuerySupported (item, NULL) == WORD7))
      {
	  wvInitFIB (item);
	  wvStream_offset (fd, -4);
	  wvGetFIB6 (item, fd);
	  return;
      }

    item->nProduct = read_16ubit (fd);
    item->lid = read_16ubit (fd);
    wvTrace (("lid is %x\n", item->lid));
    item->pnNext = (S16) read_16ubit (fd);
    temp16 = read_16ubit (fd);
    item->fDot = (temp16 & 0x0001);
    item->fGlsy = (temp16 & 0x0002) >> 1;
    item->fComplex = (temp16 & 0x0004) >> 2;
    item->fHasPic = (temp16 & 0x0008) >> 3;
    item->cQuickSaves = (temp16 & 0x00F0) >> 4;
    item->fEncrypted = (temp16 & 0x0100) >> 8;
    item->fWhichTblStm = (temp16 & 0x0200) >> 9;
    item->fReadOnlyRecommended = (temp16 & 0x0400) >> 10;
    item->fWriteReservation = (temp16 & 0x0800) >> 11;
    item->fExtChar = (temp16 & 0x1000) >> 12;
    wvTrace (("fExtChar is %d\n", item->fExtChar));
    item->fLoadOverride = (temp16 & 0x2000) >> 13;
    item->fFarEast = (temp16 & 0x4000) >> 14;
    item->fCrypto = (temp16 & 0x8000) >> 15;
    item->nFibBack = read_16ubit (fd);
    item->lKey = read_32ubit (fd);
    item->envr = read_8ubit (fd);
    temp8 = read_8ubit (fd);
    item->fMac = (temp8 & 0x01);
    item->fEmptySpecial = (temp8 & 0x02) >> 1;
    item->fLoadOverridePage = (temp8 & 0x04) >> 2;
    item->fFutureSavedUndo = (temp8 & 0x08) >> 3;
    item->fWord97Saved = (temp8 & 0x10) >> 4;
    item->fSpare0 = (temp8 & 0xFE) >> 5;
    item->chse = read_16ubit (fd);
    item->chsTables = read_16ubit (fd);
    item->fcMin = read_32ubit (fd);
    item->fcMac = read_32ubit (fd);
    item->csw = read_16ubit (fd);
    item->wMagicCreated = read_16ubit (fd);
    item->wMagicRevised = read_16ubit (fd);
    item->wMagicCreatedPrivate = read_16ubit (fd);
    item->wMagicRevisedPrivate = read_16ubit (fd);
    item->pnFbpChpFirst_W6 = (S16) read_16ubit (fd);
    item->pnChpFirst_W6 = (S16) read_16ubit (fd);
    item->cpnBteChp_W6 = (S16) read_16ubit (fd);
    item->pnFbpPapFirst_W6 = (S16) read_16ubit (fd);
    item->pnPapFirst_W6 = (S16) read_16ubit (fd);
    item->cpnBtePap_W6 = (S16) read_16ubit (fd);
    item->pnFbpLvcFirst_W6 = (S16) read_16ubit (fd);
    item->pnLvcFirst_W6 = (S16) read_16ubit (fd);
    item->cpnBteLvc_W6 = (S16) read_16ubit (fd);
    item->lidFE = (S16) read_16ubit (fd);
    item->clw = read_16ubit (fd);
    item->cbMac = (S32) read_32ubit (fd);
    item->lProductCreated = read_32ubit (fd);
    item->lProductRevised = read_32ubit (fd);
    item->ccpText = read_32ubit (fd);
    item->ccpFtn = (S32) read_32ubit (fd);
    item->ccpHdr = (S32) read_32ubit (fd);
    item->ccpMcr = (S32) read_32ubit (fd);
    item->ccpAtn = (S32) read_32ubit (fd);
    item->ccpEdn = (S32) read_32ubit (fd);
    item->ccpTxbx = (S32) read_32ubit (fd);
    item->ccpHdrTxbx = (S32) read_32ubit (fd);
    item->pnFbpChpFirst = (S32) read_32ubit (fd);
    item->pnChpFirst = (S32) read_32ubit (fd);
    item->cpnBteChp = (S32) read_32ubit (fd);
    item->pnFbpPapFirst = (S32) read_32ubit (fd);
    item->pnPapFirst = (S32) read_32ubit (fd);
    item->cpnBtePap = (S32) read_32ubit (fd);
    item->pnFbpLvcFirst = (S32) read_32ubit (fd);
    item->pnLvcFirst = (S32) read_32ubit (fd);
    item->cpnBteLvc = (S32) read_32ubit (fd);
    item->fcIslandFirst = (S32) read_32ubit (fd);
    item->fcIslandLim = (S32) read_32ubit (fd);
    item->cfclcb = read_16ubit (fd);
    item->fcStshfOrig = (S32) read_32ubit (fd);
    item->lcbStshfOrig = read_32ubit (fd);
    item->fcStshf = (S32) read_32ubit (fd);
    item->lcbStshf = read_32ubit (fd);

    item->fcPlcffndRef = (S32) read_32ubit (fd);
    item->lcbPlcffndRef = read_32ubit (fd);
    item->fcPlcffndTxt = (S32) read_32ubit (fd);
    item->lcbPlcffndTxt = read_32ubit (fd);
    item->fcPlcfandRef = (S32) read_32ubit (fd);
    item->lcbPlcfandRef = read_32ubit (fd);
    item->fcPlcfandTxt = (S32) read_32ubit (fd);
    item->lcbPlcfandTxt = read_32ubit (fd);
    item->fcPlcfsed = (S32) read_32ubit (fd);
    item->lcbPlcfsed = read_32ubit (fd);
    item->fcPlcpad = (S32) read_32ubit (fd);
    item->lcbPlcpad = read_32ubit (fd);
    item->fcPlcfphe = (S32) read_32ubit (fd);
    item->lcbPlcfphe = read_32ubit (fd);
    item->fcSttbfglsy = (S32) read_32ubit (fd);
    item->lcbSttbfglsy = read_32ubit (fd);
    item->fcPlcfglsy = (S32) read_32ubit (fd);
    item->lcbPlcfglsy = read_32ubit (fd);
    item->fcPlcfhdd = (S32) read_32ubit (fd);
    item->lcbPlcfhdd = read_32ubit (fd);
    item->fcPlcfbteChpx = (S32) read_32ubit (fd);
    item->lcbPlcfbteChpx = read_32ubit (fd);
    item->fcPlcfbtePapx = (S32) read_32ubit (fd);
    item->lcbPlcfbtePapx = read_32ubit (fd);
    item->fcPlcfsea = (S32) read_32ubit (fd);
    item->lcbPlcfsea = read_32ubit (fd);
    item->fcSttbfffn = (S32) read_32ubit (fd);
    item->lcbSttbfffn = read_32ubit (fd);
    item->fcPlcffldMom = (S32) read_32ubit (fd);
    item->lcbPlcffldMom = read_32ubit (fd);
    item->fcPlcffldHdr = (S32) read_32ubit (fd);
    item->lcbPlcffldHdr = read_32ubit (fd);
    item->fcPlcffldFtn = (S32) read_32ubit (fd);
    item->lcbPlcffldFtn = read_32ubit (fd);
    item->fcPlcffldAtn = (S32) read_32ubit (fd);
    item->lcbPlcffldAtn = read_32ubit (fd);
    item->fcPlcffldMcr = (S32) read_32ubit (fd);
    item->lcbPlcffldMcr = read_32ubit (fd);
    item->fcSttbfbkmk = (S32) read_32ubit (fd);
    item->lcbSttbfbkmk = read_32ubit (fd);
    item->fcPlcfbkf = (S32) read_32ubit (fd);
    item->lcbPlcfbkf = read_32ubit (fd);
    item->fcPlcfbkl = (S32) read_32ubit (fd);
    item->lcbPlcfbkl = read_32ubit (fd);
    item->fcCmds = (S32) read_32ubit (fd);
    item->lcbCmds = read_32ubit (fd);
    item->fcPlcmcr = (S32) read_32ubit (fd);
    item->lcbPlcmcr = read_32ubit (fd);
    item->fcSttbfmcr = (S32) read_32ubit (fd);
    item->lcbSttbfmcr = read_32ubit (fd);
    item->fcPrDrvr = (S32) read_32ubit (fd);
    item->lcbPrDrvr = read_32ubit (fd);
    item->fcPrEnvPort = (S32) read_32ubit (fd);
    item->lcbPrEnvPort = read_32ubit (fd);
    item->fcPrEnvLand = (S32) read_32ubit (fd);
    item->lcbPrEnvLand = read_32ubit (fd);
    item->fcWss = (S32) read_32ubit (fd);
    item->lcbWss = read_32ubit (fd);
    item->fcDop = (S32) read_32ubit (fd);
    item->lcbDop = read_32ubit (fd);
    item->fcSttbfAssoc = (S32) read_32ubit (fd);
    item->lcbSttbfAssoc = read_32ubit (fd);
    item->fcClx = (S32) read_32ubit (fd);
    item->lcbClx = read_32ubit (fd);
    item->fcPlcfpgdFtn = (S32) read_32ubit (fd);
    item->lcbPlcfpgdFtn = read_32ubit (fd);
    item->fcAutosaveSource = (S32) read_32ubit (fd);
    item->lcbAutosaveSource = read_32ubit (fd);
    item->fcGrpXstAtnOwners = (S32) read_32ubit (fd);
    item->lcbGrpXstAtnOwners = read_32ubit (fd);
    item->fcSttbfAtnbkmk = (S32) read_32ubit (fd);
    item->lcbSttbfAtnbkmk = read_32ubit (fd);
    item->fcPlcdoaMom = (S32) read_32ubit (fd);
    item->lcbPlcdoaMom = read_32ubit (fd);
    item->fcPlcdoaHdr = (S32) read_32ubit (fd);
    item->lcbPlcdoaHdr = read_32ubit (fd);
    item->fcPlcspaMom = (S32) read_32ubit (fd);
    item->lcbPlcspaMom = read_32ubit (fd);
    item->fcPlcspaHdr = (S32) read_32ubit (fd);
    item->lcbPlcspaHdr = read_32ubit (fd);
    item->fcPlcfAtnbkf = (S32) read_32ubit (fd);
    item->lcbPlcfAtnbkf = read_32ubit (fd);
    item->fcPlcfAtnbkl = (S32) read_32ubit (fd);
    item->lcbPlcfAtnbkl = read_32ubit (fd);
    item->fcPms = (S32) read_32ubit (fd);
    item->lcbPms = read_32ubit (fd);
    item->fcFormFldSttbs = (S32) read_32ubit (fd);
    item->lcbFormFldSttbs = read_32ubit (fd);
    item->fcPlcfendRef = (S32) read_32ubit (fd);
    item->lcbPlcfendRef = read_32ubit (fd);
    item->fcPlcfendTxt = (S32) read_32ubit (fd);
    item->lcbPlcfendTxt = read_32ubit (fd);
    item->fcPlcffldEdn = (S32) read_32ubit (fd);
    item->lcbPlcffldEdn = read_32ubit (fd);
    item->fcPlcfpgdEdn = (S32) read_32ubit (fd);
    item->lcbPlcfpgdEdn = read_32ubit (fd);
    item->fcDggInfo = (S32) read_32ubit (fd);
    item->lcbDggInfo = read_32ubit (fd);
    item->fcSttbfRMark = (S32) read_32ubit (fd);
    item->lcbSttbfRMark = read_32ubit (fd);
    item->fcSttbCaption = (S32) read_32ubit (fd);
    item->lcbSttbCaption = read_32ubit (fd);
    item->fcSttbAutoCaption = (S32) read_32ubit (fd);
    item->lcbSttbAutoCaption = read_32ubit (fd);
    item->fcPlcfwkb = (S32) read_32ubit (fd);
    item->lcbPlcfwkb = read_32ubit (fd);
    item->fcPlcfspl = (S32) read_32ubit (fd);
    item->lcbPlcfspl = read_32ubit (fd);
    item->fcPlcftxbxTxt = (S32) read_32ubit (fd);
    item->lcbPlcftxbxTxt = read_32ubit (fd);
    item->fcPlcffldTxbx = (S32) read_32ubit (fd);
    item->lcbPlcffldTxbx = read_32ubit (fd);
    item->fcPlcfhdrtxbxTxt = (S32) read_32ubit (fd);
    item->lcbPlcfhdrtxbxTxt = read_32ubit (fd);
    item->fcPlcffldHdrTxbx = (S32) read_32ubit (fd);
    item->lcbPlcffldHdrTxbx = read_32ubit (fd);
    item->fcStwUser = (S32) read_32ubit (fd);
    item->lcbStwUser = read_32ubit (fd);
    item->fcSttbttmbd = (S32) read_32ubit (fd);
    item->cbSttbttmbd = read_32ubit (fd);
    item->fcUnused = (S32) read_32ubit (fd);
    item->lcbUnused = read_32ubit (fd);
    item->fcPgdMother = (S32) read_32ubit (fd);
    item->lcbPgdMother = read_32ubit (fd);
    item->fcBkdMother = (S32) read_32ubit (fd);
    item->lcbBkdMother = read_32ubit (fd);
    item->fcPgdFtn = (S32) read_32ubit (fd);
    item->lcbPgdFtn = read_32ubit (fd);
    item->fcBkdFtn = (S32) read_32ubit (fd);
    item->lcbBkdFtn = read_32ubit (fd);
    item->fcPgdEdn = (S32) read_32ubit (fd);
    item->lcbPgdEdn = read_32ubit (fd);
    item->fcBkdEdn = (S32) read_32ubit (fd);
    item->lcbBkdEdn = read_32ubit (fd);
    item->fcSttbfIntlFld = (S32) read_32ubit (fd);
    item->lcbSttbfIntlFld = read_32ubit (fd);
    item->fcRouteSlip = (S32) read_32ubit (fd);
    item->lcbRouteSlip = read_32ubit (fd);
    item->fcSttbSavedBy = (S32) read_32ubit (fd);
    item->lcbSttbSavedBy = read_32ubit (fd);
    item->fcSttbFnm = (S32) read_32ubit (fd);
    item->lcbSttbFnm = read_32ubit (fd);
    item->fcPlcfLst = (S32) read_32ubit (fd);
    item->lcbPlcfLst = read_32ubit (fd);
    item->fcPlfLfo = (S32) read_32ubit (fd);
    item->lcbPlfLfo = read_32ubit (fd);
    item->fcPlcftxbxBkd = (S32) read_32ubit (fd);
    item->lcbPlcftxbxBkd = read_32ubit (fd);
    item->fcPlcftxbxHdrBkd = (S32) read_32ubit (fd);
    item->lcbPlcftxbxHdrBkd = read_32ubit (fd);
    item->fcDocUndo = (S32) read_32ubit (fd);
    item->lcbDocUndo = read_32ubit (fd);
    item->fcRgbuse = (S32) read_32ubit (fd);
    item->lcbRgbuse = read_32ubit (fd);
    item->fcUsp = (S32) read_32ubit (fd);
    item->lcbUsp = read_32ubit (fd);
    item->fcUskf = (S32) read_32ubit (fd);
    item->lcbUskf = read_32ubit (fd);
    item->fcPlcupcRgbuse = (S32) read_32ubit (fd);
    item->lcbPlcupcRgbuse = read_32ubit (fd);
    item->fcPlcupcUsp = (S32) read_32ubit (fd);
    item->lcbPlcupcUsp = read_32ubit (fd);
    item->fcSttbGlsyStyle = (S32) read_32ubit (fd);
    item->lcbSttbGlsyStyle = read_32ubit (fd);
    item->fcPlgosl = (S32) read_32ubit (fd);
    item->lcbPlgosl = read_32ubit (fd);
    item->fcPlcocx = (S32) read_32ubit (fd);
    item->lcbPlcocx = read_32ubit (fd);
    item->fcPlcfbteLvc = (S32) read_32ubit (fd);
    item->lcbPlcfbteLvc = read_32ubit (fd);
    wvGetFILETIME (&(item->ftModified), fd);
    item->fcPlcflvc = (S32) read_32ubit (fd);
    item->lcbPlcflvc = read_32ubit (fd);
    item->fcPlcasumy = (S32) read_32ubit (fd);
    item->lcbPlcasumy = read_32ubit (fd);
    item->fcPlcfgram = (S32) read_32ubit (fd);
    item->lcbPlcfgram = read_32ubit (fd);
    item->fcSttbListNames = (S32) read_32ubit (fd);
    item->lcbSttbListNames = read_32ubit (fd);
    item->fcSttbfUssr = (S32) read_32ubit (fd);
    item->lcbSttbfUssr = read_32ubit (fd);
}

wvStream *
wvWhichTableStream (FIB * fib, wvParseStruct * ps)
{
    wvStream *ret;

    if ((wvQuerySupported (fib, NULL) & 0x7fff) == WORD8)
      {
	  if (fib->fWhichTblStm)
	    {
		wvTrace (("1Table\n"));
		ret = ps->tablefd1;
		if (ret == NULL)
		  {
		      wvError (
			       ("!!, the FIB lied to us, (told us to use the 1Table) making a heroic effort to use the other table stream, hold on tight\n"));
		      ret = ps->tablefd0;
		  }
	    }
	  else
	    {
		wvTrace (("0Table\n"));
		ret = ps->tablefd0;
		if (ret == NULL)
		  {
		      wvError (
			       ("!!, the FIB lied to us, (told us to use the 0Table) making a heroic effort to use the other table stream, hold on tight\n"));
		      ret = ps->tablefd1;
		  }
	    }
      }
    else			/* word 7- */
	ret = ps->mainfd;
    return (ret);
}


wvVersion
wvQuerySupported (FIB * fib, int *reason)
{
    int ret = WORD8;

    if (fib->wIdent == 0x37FE)
	ret = WORD5;
    else
      {
	  /*begin from microsofts kb q 40 */
	  if (fib->nFib < 101)
	    {
		if (reason)
		    *reason = 1;
		ret = WORD2;
	    }
	  else
	    {
		switch (fib->nFib)
		  {
		  case 101:
		      if (reason)
			  *reason = 2;
		      ret = WORD6;
		      break;	/* I'm pretty sure we should break here, Jamie. */
		  case 103:
		  case 104:
		      if (reason)
			  *reason = 3;
		      ret = WORD7;
		      break;	/* I'm pretty sure we should break here, Jamie. */
		  default:
		      break;
		  }
	    }
	  /*end from microsofts kb q 40 */
      }
    wvTrace (("RET is %d\n", ret));
    if (fib->fEncrypted)
      {
	  if (reason)
	      *reason = 4;
	  ret |= 0x8000;
      }
    return (ret);
}

void
wvGetFIB2 (FIB * item, wvStream * fd)
{
    U16 temp16 = 0;

    item->wIdent = read_16ubit (fd);
    item->nFib = read_16ubit (fd);

    item->nProduct = read_16ubit (fd);
    item->lid = read_16ubit (fd);
    wvTrace (("lid is %x\n", item->lid));
    item->pnNext = (S16) read_16ubit (fd);
    temp16 = read_16ubit (fd);

    item->fDot = (temp16 & 0x0001);
    item->fGlsy = (temp16 & 0x0002) >> 1;
    item->fComplex = (temp16 & 0x0004) >> 2;
    item->fHasPic = (temp16 & 0x0008) >> 3;
    item->cQuickSaves = (temp16 & 0x00F0) >> 4;
    item->fEncrypted = (temp16 & 0x0100) >> 8;
    item->fWhichTblStm = 0;	/* Unused from here on */
    item->fReadOnlyRecommended = 0;
    item->fWriteReservation = 0;
    item->fExtChar = 0;
    item->fLoadOverride = 0;
    item->fFarEast = 0;
    item->fCrypto = 0;

    item->nFibBack = read_16ubit (fd);
    wvTrace (("nFibBack is %d\n", item->nFibBack));

    item->Spare = read_32ubit (fd);	/* A spare for W2 */
    item->rgwSpare0[0] = read_16ubit (fd);
    item->rgwSpare0[1] = read_16ubit (fd);
    item->rgwSpare0[2] = read_16ubit (fd);
    item->fcMin = read_32ubit (fd);	/* These appear correct MV 29.8.2000 */
    item->fcMac = read_32ubit (fd);
    wvTrace (("fc from %d to %d\n", item->fcMin, item->fcMac));

    item->cbMac = read_32ubit (fd);	/* Last byte file position plus one. */

    item->fcSpare0 = read_32ubit (fd);
    item->fcSpare1 = read_32ubit (fd);
    item->fcSpare2 = read_32ubit (fd);
    item->fcSpare3 = read_32ubit (fd);

    item->ccpText = read_32ubit (fd);
    wvTrace (("length %d == %d\n", item->fcMac - item->fcMin, item->ccpText));

    item->ccpFtn = (S32) read_32ubit (fd);
    item->ccpHdr = (S32) read_32ubit (fd);
    item->ccpMcr = (S32) read_32ubit (fd);
    item->ccpAtn = (S32) read_32ubit (fd);
    item->ccpSpare0 = (S32) read_32ubit (fd);
    item->ccpSpare1 = (S32) read_32ubit (fd);
    item->ccpSpare2 = (S32) read_32ubit (fd);
    item->ccpSpare3 = (S32) read_32ubit (fd);

    item->fcStshfOrig = read_32ubit (fd);
    item->lcbStshfOrig = (S32) read_16ubit (fd);
    item->fcStshf = read_32ubit (fd);
    item->lcbStshf = (S32) read_16ubit (fd);
    item->fcPlcffndRef = read_32ubit (fd);
    item->lcbPlcffndRef = (S32) read_16ubit (fd);
    item->fcPlcffndTxt = read_32ubit (fd);
    item->lcbPlcffndTxt = (S32) read_16ubit (fd);
    item->fcPlcfandRef = read_32ubit (fd);
    item->lcbPlcfandRef = (S32) read_16ubit (fd);
    item->fcPlcfandTxt = read_32ubit (fd);
    item->lcbPlcfandTxt = (S32) read_16ubit (fd);
    item->fcPlcfsed = read_32ubit (fd);
    item->lcbPlcfsed = (S32) read_16ubit (fd);
    item->fcPlcfpgd = read_32ubit (fd);
    item->cbPlcfpgd = read_16ubit (fd);
    item->fcPlcfphe = read_32ubit (fd);
    item->lcbPlcfphe = (S32) read_16ubit (fd);
    item->fcPlcfglsy = read_32ubit (fd);
    item->lcbPlcfglsy = (S32) read_16ubit (fd);
    item->fcPlcfhdd = read_32ubit (fd);
    item->lcbPlcfhdd = (S32) read_16ubit (fd);
    item->fcPlcfbteChpx = read_32ubit (fd);
    item->lcbPlcfbteChpx = (S32) read_16ubit (fd);
    item->fcPlcfbtePapx = read_32ubit (fd);
    item->lcbPlcfbtePapx = (S32) read_16ubit (fd);
    item->fcPlcfsea = read_32ubit (fd);
    item->lcbPlcfsea = (S32) read_16ubit (fd);
    item->fcSttbfffn = read_32ubit (fd);
    item->lcbSttbfffn = (S32) read_16ubit (fd);
    item->fcPlcffldMom = read_32ubit (fd);
    item->lcbPlcffldMom = (S32) read_16ubit (fd);
    item->fcPlcffldHdr = read_32ubit (fd);
    item->lcbPlcffldHdr = (S32) read_16ubit (fd);
    item->fcPlcffldFtn = read_32ubit (fd);
    item->lcbPlcffldFtn = (S32) read_16ubit (fd);
    item->fcPlcffldAtn = read_32ubit (fd);
    item->lcbPlcffldAtn = (S32) read_16ubit (fd);
    item->fcPlcffldMcr = read_32ubit (fd);
    item->lcbPlcffldMcr = (S32) read_16ubit (fd);
    item->fcSttbfbkmk = read_32ubit (fd);
    item->lcbSttbfbkmk = (S32) read_16ubit (fd);
    item->fcPlcfbkf = read_32ubit (fd);
    item->lcbPlcfbkf = (S32) read_16ubit (fd);
    item->fcPlcfbkl = read_32ubit (fd);
    item->lcbPlcfbkl = (S32) read_16ubit (fd);
    item->fcCmds = read_32ubit (fd);
    item->lcbCmds = (S32) read_16ubit (fd);
    item->fcPlcmcr = read_32ubit (fd);
    item->lcbPlcmcr = (S32) read_16ubit (fd);
    item->fcSttbfmcr = read_32ubit (fd);
    item->lcbSttbfmcr = (S32) read_16ubit (fd);
    item->fcPrDrvr = read_32ubit (fd);
    item->lcbPrDrvr = (S32) read_16ubit (fd);
    item->fcPrEnvPort = read_32ubit (fd);
    item->lcbPrEnvPort = (S32) read_16ubit (fd);
    item->fcPrEnvLand = read_32ubit (fd);
    item->lcbPrEnvLand = (S32) read_16ubit (fd);
    item->fcWss = read_32ubit (fd);
    item->lcbWss = (S32) read_16ubit (fd);
    item->fcDop = read_32ubit (fd);
    item->lcbDop = (S32) read_16ubit (fd);
    item->fcSttbfAssoc = read_32ubit (fd);
    item->lcbSttbfAssoc = (S32) read_16ubit (fd);
    item->fcClx = read_32ubit (fd);
    item->lcbClx = (S32) read_16ubit (fd);
    item->fcPlcfpgdFtn = read_32ubit (fd);
    item->lcbPlcfpgdFtn = (S32) read_16ubit (fd);
    item->fcAutosaveSource = read_32ubit (fd);
    item->lcbAutosaveSource = (S32) read_16ubit (fd);
    item->fcSpare5 = read_32ubit (fd);
    item->cbSpare5 = read_16ubit (fd);
    item->fcSpare6 = read_32ubit (fd);
    item->cbSpare6 = read_16ubit (fd);
    item->wSpare4 = read_16ubit (fd);
    item->pnChpFirst = read_16ubit (fd);
    item->pnPapFirst = read_16ubit (fd);
    item->cpnBteChp = read_16ubit (fd);
    item->cpnBtePap = read_16ubit (fd);

}

void
wvGetFIB6 (FIB * item, wvStream * fd)
{
    U16 temp16;
    U8 temp8;

    item->wIdent = read_16ubit (fd);
    item->nFib = read_16ubit (fd);

    item->nProduct = read_16ubit (fd);
    item->lid = read_16ubit (fd);
    wvTrace (("lid is %x\n", item->lid));
    item->pnNext = (S16) read_16ubit (fd);
    temp16 = read_16ubit (fd);

    item->fDot = (temp16 & 0x0001);
    item->fGlsy = (temp16 & 0x0002) >> 1;
    item->fComplex = (temp16 & 0x0004) >> 2;
    item->fHasPic = (temp16 & 0x0008) >> 3;
    item->cQuickSaves = (temp16 & 0x00F0) >> 4;
    item->fEncrypted = (temp16 & 0x0100) >> 8;
    item->fWhichTblStm = 0;	/* word 6 files only have one table stream */
    item->fReadOnlyRecommended = (temp16 & 0x0400) >> 10;
    item->fWriteReservation = (temp16 & 0x0800) >> 11;
    item->fExtChar = (temp16 & 0x1000) >> 12;
    wvTrace (("fExtChar is %d\n", item->fExtChar));
    item->fLoadOverride = 0;
    item->fFarEast = 0;
    item->fCrypto = 0;
    item->nFibBack = read_16ubit (fd);
    item->lKey = read_32ubit (fd);
    item->envr = read_8ubit (fd);
    temp8 = read_8ubit (fd);
    item->fMac = 0;
    item->fEmptySpecial = 0;
    item->fLoadOverridePage = 0;
    item->fFutureSavedUndo = 0;
    item->fWord97Saved = 0;
    item->fSpare0 = 0;
    item->chse = read_16ubit (fd);
    item->chsTables = read_16ubit (fd);
    item->fcMin = read_32ubit (fd);
    item->fcMac = read_32ubit (fd);

    item->csw = 14;
    item->wMagicCreated = 0xCA0;	/*this is the unique id of the creater, so its me :-) */

    item->cbMac = read_32ubit (fd);

    read_16ubit (fd);
    read_16ubit (fd);
    read_16ubit (fd);
    read_16ubit (fd);
    read_16ubit (fd);
    read_16ubit (fd);
    read_16ubit (fd);
    read_16ubit (fd);

    item->ccpText = read_32ubit (fd);
    item->ccpFtn = (S32) read_32ubit (fd);
    item->ccpHdr = (S32) read_32ubit (fd);
    item->ccpMcr = (S32) read_32ubit (fd);
    item->ccpAtn = (S32) read_32ubit (fd);
    item->ccpEdn = (S32) read_32ubit (fd);
    item->ccpTxbx = (S32) read_32ubit (fd);
    item->ccpHdrTxbx = (S32) read_32ubit (fd);

    read_32ubit (fd);

    item->fcStshfOrig = (S32) read_32ubit (fd);
    item->lcbStshfOrig = read_32ubit (fd);
    item->fcStshf = (S32) read_32ubit (fd);
    item->lcbStshf = read_32ubit (fd);
    item->fcPlcffndRef = (S32) read_32ubit (fd);
    item->lcbPlcffndRef = read_32ubit (fd);
    item->fcPlcffndTxt = (S32) read_32ubit (fd);
    item->lcbPlcffndTxt = read_32ubit (fd);
    item->fcPlcfandRef = (S32) read_32ubit (fd);
    item->lcbPlcfandRef = read_32ubit (fd);
    item->fcPlcfandTxt = (S32) read_32ubit (fd);
    item->lcbPlcfandTxt = read_32ubit (fd);
    item->fcPlcfsed = (S32) read_32ubit (fd);
    item->lcbPlcfsed = read_32ubit (fd);
    item->fcPlcpad = (S32) read_32ubit (fd);
    item->lcbPlcpad = read_32ubit (fd);
    item->fcPlcfphe = (S32) read_32ubit (fd);
    item->lcbPlcfphe = read_32ubit (fd);
    item->fcSttbfglsy = (S32) read_32ubit (fd);
    item->lcbSttbfglsy = read_32ubit (fd);
    item->fcPlcfglsy = (S32) read_32ubit (fd);
    item->lcbPlcfglsy = read_32ubit (fd);
    item->fcPlcfhdd = (S32) read_32ubit (fd);
    item->lcbPlcfhdd = read_32ubit (fd);
    item->fcPlcfbteChpx = (S32) read_32ubit (fd);
    item->lcbPlcfbteChpx = read_32ubit (fd);
    item->fcPlcfbtePapx = (S32) read_32ubit (fd);
    item->lcbPlcfbtePapx = read_32ubit (fd);
    item->fcPlcfsea = (S32) read_32ubit (fd);
    item->lcbPlcfsea = read_32ubit (fd);
    item->fcSttbfffn = (S32) read_32ubit (fd);
    item->lcbSttbfffn = read_32ubit (fd);
    item->fcPlcffldMom = (S32) read_32ubit (fd);
    item->lcbPlcffldMom = read_32ubit (fd);
    item->fcPlcffldHdr = (S32) read_32ubit (fd);
    item->lcbPlcffldHdr = read_32ubit (fd);
    item->fcPlcffldFtn = (S32) read_32ubit (fd);
    item->lcbPlcffldFtn = read_32ubit (fd);
    item->fcPlcffldAtn = (S32) read_32ubit (fd);
    item->lcbPlcffldAtn = read_32ubit (fd);
    item->fcPlcffldMcr = (S32) read_32ubit (fd);
    item->lcbPlcffldMcr = read_32ubit (fd);
    item->fcSttbfbkmk = (S32) read_32ubit (fd);
    item->lcbSttbfbkmk = read_32ubit (fd);
    item->fcPlcfbkf = (S32) read_32ubit (fd);
    item->lcbPlcfbkf = read_32ubit (fd);
    item->fcPlcfbkl = (S32) read_32ubit (fd);
    item->lcbPlcfbkl = read_32ubit (fd);
    item->fcCmds = (S32) read_32ubit (fd);
    item->lcbCmds = read_32ubit (fd);
    item->fcPlcmcr = (S32) read_32ubit (fd);
    item->lcbPlcmcr = read_32ubit (fd);
    item->fcSttbfmcr = (S32) read_32ubit (fd);
    item->lcbSttbfmcr = read_32ubit (fd);
    item->fcPrDrvr = (S32) read_32ubit (fd);
    item->lcbPrDrvr = read_32ubit (fd);
    item->fcPrEnvPort = (S32) read_32ubit (fd);
    item->lcbPrEnvPort = read_32ubit (fd);
    item->fcPrEnvLand = (S32) read_32ubit (fd);
    item->lcbPrEnvLand = read_32ubit (fd);
    item->fcWss = (S32) read_32ubit (fd);
    item->lcbWss = read_32ubit (fd);
    item->fcDop = (S32) read_32ubit (fd);
    item->lcbDop = read_32ubit (fd);
    item->fcSttbfAssoc = (S32) read_32ubit (fd);
    item->lcbSttbfAssoc = read_32ubit (fd);
    item->fcClx = (S32) read_32ubit (fd);
    item->lcbClx = read_32ubit (fd);
    item->fcPlcfpgdFtn = (S32) read_32ubit (fd);
    item->lcbPlcfpgdFtn = read_32ubit (fd);
    item->fcAutosaveSource = (S32) read_32ubit (fd);
    item->lcbAutosaveSource = read_32ubit (fd);
    item->fcGrpXstAtnOwners = (S32) read_32ubit (fd);
    item->lcbGrpXstAtnOwners = read_32ubit (fd);
    item->fcSttbfAtnbkmk = (S32) read_32ubit (fd);
    item->lcbSttbfAtnbkmk = read_32ubit (fd);

    read_16ubit (fd);

    item->pnChpFirst = (S32) read_16ubit (fd);
    item->pnPapFirst = (S32) read_16ubit (fd);
    item->cpnBteChp = (S32) read_16ubit (fd);
    item->cpnBtePap = (S32) read_16ubit (fd);
    item->fcPlcdoaMom = (S32) read_32ubit (fd);
    item->lcbPlcdoaMom = read_32ubit (fd);
    item->fcPlcdoaHdr = (S32) read_32ubit (fd);
    item->lcbPlcdoaHdr = read_32ubit (fd);

    read_32ubit (fd);
    read_32ubit (fd);
    read_32ubit (fd);
    read_32ubit (fd);

    item->fcPlcfAtnbkf = (S32) read_32ubit (fd);
    item->lcbPlcfAtnbkf = read_32ubit (fd);
    item->fcPlcfAtnbkl = (S32) read_32ubit (fd);
    item->lcbPlcfAtnbkl = read_32ubit (fd);
    item->fcPms = (S32) read_32ubit (fd);
    item->lcbPms = read_32ubit (fd);
    item->fcFormFldSttbs = (S32) read_32ubit (fd);
    item->lcbFormFldSttbs = read_32ubit (fd);
    item->fcPlcfendRef = (S32) read_32ubit (fd);
    item->lcbPlcfendRef = read_32ubit (fd);
    item->fcPlcfendTxt = (S32) read_32ubit (fd);
    item->lcbPlcfendTxt = read_32ubit (fd);
    item->fcPlcffldEdn = (S32) read_32ubit (fd);
    item->lcbPlcffldEdn = read_32ubit (fd);
    item->fcPlcfpgdEdn = (S32) read_32ubit (fd);
    item->lcbPlcfpgdEdn = read_32ubit (fd);

    read_32ubit (fd);
    read_32ubit (fd);

    item->fcSttbfRMark = (S32) read_32ubit (fd);
    item->lcbSttbfRMark = read_32ubit (fd);
    item->fcSttbCaption = (S32) read_32ubit (fd);
    item->lcbSttbCaption = read_32ubit (fd);
    item->fcSttbAutoCaption = (S32) read_32ubit (fd);
    item->lcbSttbAutoCaption = read_32ubit (fd);
    item->fcPlcfwkb = (S32) read_32ubit (fd);
    item->lcbPlcfwkb = read_32ubit (fd);

    read_32ubit (fd);
    read_32ubit (fd);


    item->fcPlcftxbxTxt = (S32) read_32ubit (fd);
    item->lcbPlcftxbxTxt = read_32ubit (fd);
    item->fcPlcffldTxbx = (S32) read_32ubit (fd);
    item->lcbPlcffldTxbx = read_32ubit (fd);
    item->fcPlcfhdrtxbxTxt = (S32) read_32ubit (fd);
    item->lcbPlcfhdrtxbxTxt = read_32ubit (fd);
    item->fcPlcffldHdrTxbx = (S32) read_32ubit (fd);
    item->lcbPlcffldHdrTxbx = read_32ubit (fd);
    item->fcStwUser = (S32) read_32ubit (fd);
    item->lcbStwUser = read_32ubit (fd);
    item->fcSttbttmbd = (S32) read_32ubit (fd);
    item->cbSttbttmbd = read_32ubit (fd);
    item->fcUnused = (S32) read_32ubit (fd);
    item->lcbUnused = read_32ubit (fd);
    item->fcPgdMother = (S32) read_32ubit (fd);
    item->lcbPgdMother = read_32ubit (fd);
    item->fcBkdMother = (S32) read_32ubit (fd);
    item->lcbBkdMother = read_32ubit (fd);
    item->fcPgdFtn = (S32) read_32ubit (fd);
    item->lcbPgdFtn = read_32ubit (fd);
    item->fcBkdFtn = (S32) read_32ubit (fd);
    item->lcbBkdFtn = read_32ubit (fd);
    item->fcPgdEdn = (S32) read_32ubit (fd);
    item->lcbPgdEdn = read_32ubit (fd);
    item->fcBkdEdn = (S32) read_32ubit (fd);
    item->lcbBkdEdn = read_32ubit (fd);
    item->fcSttbfIntlFld = (S32) read_32ubit (fd);
    item->lcbSttbfIntlFld = read_32ubit (fd);
    item->fcRouteSlip = (S32) read_32ubit (fd);
    item->lcbRouteSlip = read_32ubit (fd);
    item->fcSttbSavedBy = (S32) read_32ubit (fd);
    item->lcbSttbSavedBy = read_32ubit (fd);
    item->fcSttbFnm = (S32) read_32ubit (fd);
    item->lcbSttbFnm = read_32ubit (fd);
}

```


## decode_complex.c
```c
/* wvWare
 * Copyright (C) Caolan McNamara, Dom Lachowicz, and others
 *
 * This program is free software; you can redistribute it and/or
 * modify it under the terms of the GNU General Public License
 * as published by the Free Software Foundation; either version 2
 * of the License, or (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program; if not, write to the Free Software
 * Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA
 * 02111-1307, USA.
 */

#ifdef HAVE_CONFIG_H
#include "config.h"
#endif

#include <stdlib.h>
#include <stdio.h>
#include "wv.h"

/*
To find the beginning of the paragraph containing a character in a complex
document, it's first necessary to 

1) search for the piece containing the character in the piece table. 

2) Then calculate the FC in the file that stores the character from the piece 
	table information. 
	
3) Using the FC, search the FCs FKP for the largest FC less than the character's 
	FC, call it fcTest. 
	
4) If the character at fcTest-1 is contained in the current piece, then the 
	character corresponding to that FC in the piece is the first character of 
	the paragraph. 
	
5) If that FC is before or marks the beginning of the piece, scan a piece at a 
time towards the beginning of the piece table until a piece is found that 
contains a paragraph mark. 

(This can be done by using the end of the piece FC, finding the largest FC in 
its FKP that is less than or equal to the end of piece FC, and checking to see 
if the character in front of the FKP FC (which must mark a paragraph end) is 
within the piece.)

6) When such an FKP FC is found, the FC marks the first byte of paragraph text.
*/

/*
To find the end of a paragraph for a character in a complex format file,
again 

1) it is necessary to know the piece that contains the character and the
FC assigned to the character. 

2) Using the FC of the character, first search the FKP that describes the 
character to find the smallest FC in the rgfc that is larger than the character 
FC. 

3) If the FC found in the FKP is less than or equal to the limit FC of the 
piece, the end of the paragraph that contains the character is at the FKP FC 
minus 1. 

4) If the FKP FC that was found was greater than the FC of the end of the 
piece, scan piece by piece toward the end of the document until a piece is 
found that contains a paragraph end mark. 

5) It's possible to check if a piece contains a paragraph mark by using the 
FC of the beginning of the piece to search in the FKPs for the smallest FC in 
the FKP rgfc that is greater than the FC of the beginning of the piece. 

If the FC found is less than or equal to the limit FC of the
piece, then the character that ends the paragraph is the character
immediately before the FKP FC.
*/
int
wvGetComplexParaBounds (wvVersion ver, PAPX_FKP * fkp, U32 * fcFirst,
			U32 * fcLim, U32 currentfc, CLX * clx, BTE * bte,
			U32 * pos, int nobte, U32 piece, wvStream * fd)
{
    /*
       U32 currentfc;
     */
    BTE entry;
    long currentpos;

    if (currentfc == 0xffffffffL)
      {
	  wvError (
		   ("Para Bounds not found !, this is ok if this is the last para, otherwise its a disaster\n"));
	  return (-1);
      }

    if (0 != wvGetBTE_FromFC (&entry, currentfc, bte, pos, nobte))
      {
	  wvError (("BTE not found !\n"));
	  return (-1);
      }
    currentpos = wvStream_tell (fd);
    /*The pagenumber of the FKP is entry.pn */

    wvTrace (("the entry.pn is %d\n", entry.pn));
    wvGetPAPX_FKP (ver, fkp, entry.pn, fd);

    wvGetComplexParafcFirst (ver, fcFirst, currentfc, clx, bte, pos, nobte,
			     piece, fkp, fd);

    wvReleasePAPX_FKP (fkp);
    wvTrace (("BREAK\n"));
    wvGetPAPX_FKP (ver, fkp, entry.pn, fd);

    piece =
	wvGetComplexParafcLim (ver, fcLim, currentfc, clx, bte, pos, nobte,
			       piece, fkp, fd);

    wvStream_goto (fd, currentpos);
    return (piece);
}

int
wvGetComplexParafcLim (wvVersion ver, U32 * fcLim, U32 currentfc, CLX * clx,
		       BTE * bte, U32 * pos, int nobte, U32 piece,
		       PAPX_FKP * fkp, wvStream * fd)
{
    U32 fcTest, beginfc;
    BTE entry;
    *fcLim = 0xffffffffL;
    wvTrace (("here is fcLim, currentfc is %x\n", currentfc));
    fcTest = wvSearchNextSmallestFCPAPX_FKP (fkp, currentfc);

    wvTrace (
	     ("fcTest is %x, end is %x\n", fcTest,
	      wvGetEndFCPiece (piece, clx)));


    if (fcTest <= wvGetEndFCPiece (piece, clx))
      {
	  *fcLim = fcTest;
      }
    else
      {
	  /*get end fc of previous piece */
	  piece++;
	  while (piece < clx->nopcd)
	    {
		wvTrace (("piece is %d\n", piece));
		beginfc = wvNormFC (clx->pcd[piece].fc, NULL);
		if (0 != wvGetBTE_FromFC (&entry, beginfc, bte, pos, nobte))
		  {
		      wvError (("BTE not found !\n"));
		      return (-1);
		  }
		wvReleasePAPX_FKP (fkp);
		wvGetPAPX_FKP (ver, fkp, entry.pn, fd);
		fcTest = wvSearchNextSmallestFCPAPX_FKP (fkp, beginfc);
		wvTrace (
			 ("fcTest(t) is %x, end is %x\n", fcTest,
			  wvGetEndFCPiece (piece, clx)));
		if (fcTest <= wvGetEndFCPiece (piece, clx))
		  {
		      *fcLim = fcTest;
		      break;
		  }
		piece++;
	    }
      }
    wvTrace (("fcLim is %x\n", *fcLim));
    if (piece == clx->nopcd)
      {
	  wvTrace (("failed to find a solution to end of paragraph\n"));
	  *fcLim = fcTest;
	  return (clx->nopcd - 1);	/* test using this */
      }
    return (piece);
}


int
wvGetComplexParafcFirst (wvVersion ver, U32 * fcFirst, U32 currentfc,
			 CLX * clx, BTE * bte, U32 * pos, int nobte,
			 U32 piece, PAPX_FKP * fkp, wvStream * fd)
{
    U32 fcTest, endfc;
    BTE entry;
    fcTest = wvSearchNextLargestFCPAPX_FKP (fkp, currentfc);

    wvTrace (("fcTest (s) is %x\n", fcTest));

    if (wvQuerySamePiece (fcTest - 1, clx, piece))
      {
	  wvTrace (("same piece\n"));
	  *fcFirst = fcTest - 1;
      }
    else
      {
	  /*
	     get end fc of previous piece ??, or use the end of the current piece
	   */
	  piece--;
	  while (piece != 0xffffffffL)
	    {
		wvTrace (("piece is %d\n", piece));
		endfc = wvGetEndFCPiece (piece, clx);
		wvTrace (("endfc is %x\n", endfc));
		if (0 != wvGetBTE_FromFC (&entry, endfc, bte, pos, nobte))
		  {
		      wvError (("BTE not found !\n"));
		      return (-1);
		  }
		wvReleasePAPX_FKP (fkp);
		wvGetPAPX_FKP (ver, fkp, entry.pn, fd);
		fcTest = wvSearchNextLargestFCPAPX_FKP (fkp, endfc);
		wvTrace (("fcTest(ft) is %x\n", fcTest));
		if (wvQuerySamePiece (fcTest - 1, clx, piece))
		  {
		      *fcFirst = fcTest - 1;
		      break;
		  }
		piece--;
	    }

      }
    if (piece == 0xffffffffL)
      {
	  wvTrace (
		   ("failed to find a solution to the beginning of the paragraph\n"));
	  *fcFirst = currentfc;
      }
    wvTrace (("fcFirst is finally %x\n", *fcFirst));
    return (0);
}


/* char properties version of the above -JB */
/* only difference is that we're using CHPX FKP pages,
 * and specifically just the Get and Release functions are
 * different between the two. We might be able to 
 * abstract the necessary functions to avoid duplicating them... */

int
wvGetComplexCharBounds (wvVersion ver, CHPX_FKP * fkp, U32 * fcFirst,
			U32 * fcLim, U32 currentfc, CLX * clx, BTE * bte,
			U32 * pos, int nobte, U32 piece, wvStream * fd)
{
    BTE entry;
    long currentpos;

    wvTrace (("current fc is %x\n", currentfc));

    if (currentfc == 0xffffffffL)
      {
	  wvTrace (
		   ("Char Bounds not found !, this is ok if this is the last char, otherwise its a disaster\n"));
	  return (-1);
      }

    if (0 != wvGetBTE_FromFC (&entry, currentfc, bte, pos, nobte))
      {
	  wvError (("BTE not found !\n"));
	  return (-1);
      }
    currentpos = wvStream_tell (fd);
    /*The pagenumber of the FKP is entry.pn */

    wvGetCHPX_FKP (ver, fkp, entry.pn, fd);

    wvGetComplexCharfcFirst (ver, fcFirst, currentfc, clx, bte, pos, nobte,
			     piece, fkp, fd);
    wvTrace (("BEFORE PIECE is %d\n", piece));

    wvReleaseCHPX_FKP (fkp);
    wvGetCHPX_FKP (ver, fkp, entry.pn, fd);

    piece =
	wvGetComplexCharfcLim (ver, fcLim, currentfc, clx, bte, pos, nobte,
			       piece, fkp, fd);
    wvTrace (("AFTER PIECE is %d\n", piece));

    wvStream_goto (fd, currentpos);
    return (piece);
}

int
wvGetComplexCharfcLim (wvVersion ver, U32 * fcLim, U32 currentfc, CLX * clx,
		       BTE * bte, U32 * pos, int nobte, U32 piece,
		       CHPX_FKP * fkp, wvStream * fd)
{
    U32 fcTest;
    /*
       BTE entry;
     */
    *fcLim = 0xffffffffL;
    /* this only works with the initial rgfc array, which is the
     * same for both CHPX and PAPX FKPs */
    fcTest = wvSearchNextSmallestFCPAPX_FKP ((PAPX_FKP *) fkp, currentfc);

    wvTrace (("fcTest is %x\n", fcTest));

    /*
       this single line replaces all the rest, is it conceivable that i overengineered,
       careful rereading of the spec makes no mention of repeating the para process to
       find the boundaries of the exception text runs
     */
    *fcLim = fcTest;
    wvTrace (("fcLim is %x\n", *fcLim));
    if (piece == clx->nopcd)
	return (clx->nopcd - 1);	/* test using this */
    return (piece);
}


int
wvGetComplexCharfcFirst (wvVersion ver, U32 * fcFirst, U32 currentfc,
			 CLX * clx, BTE * bte, U32 * pos, int nobte,
			 U32 piece, CHPX_FKP * fkp, wvStream * fd)
{
    U32 fcTest /*,endfc */ ;
    /*BTE entry; */
    /* this only works with the initial rgfc array, which is the */
    fcTest = wvSearchNextLargestFCCHPX_FKP (fkp, currentfc);

    wvTrace (("fcTest (s) is %x\n", fcTest));

    /*
       this single line replaces all the rest, is it conceivable that i overengineered,
       careful rereading of the spec makes no mention of repeating the para process to
       find the boundaries of the exception text runs
     */
    *fcFirst = fcTest;
    return (0);
}

/*
how this works,
we seek to the beginning of the text, we loop for a count of charaters that is stored in the fib.

the piecetable divides the text up into various sections, we keep track of our location vs
the next entry in that table, when we reach that location, we seek to the position that
the table tells us to go.

there are special cases for coming to the end of a section, and for the beginning and ends of
pages. for the purposes of headers and footers etc.
*/
void
wvDecodeComplex (wvParseStruct * ps)
{
    U32 piececount = 0, i, j, spiece = 0;
    U32 beginfc, endfc;
	U32 stream_size;
    U32 begincp, endcp;
    int ichartype;
    U8  chartype;
    U16 eachchar;
    U32 para_fcFirst, para_fcLim = 0xffffffffL;
    U32 dummy, nextpara_fcLim = 0xffffffffL;
    U32 char_fcFirst, char_fcLim = 0xffffffffL;
    U32 section_fcFirst, section_fcLim = 0xffffffffL;
    U32 comment_cpFirst = 0xffffffffL, comment_cpLim = 0xffffffffL;
    BTE *btePapx = NULL, *bteChpx = NULL;
    U32 *posPapx = NULL, *posChpx = NULL;
    U32 para_intervals, char_intervals, section_intervals, atrd_intervals;
    int cpiece = 0, npiece = 0;
    PAPX_FKP para_fkp;
    PAP apap;
    CHPX_FKP char_fkp;
    CHP achp;
    int para_pendingclose = 0, comment_pendingclose = 0, char_pendingclose =
	0, section_pendingclose = 0;
    int para_dirty = 0, char_dirty = 0, section_dirty = 0;
    SED *sed;
    SEP sep;
    U32 *posSedx;
    ATRD *atrd, *catrd = NULL;
    U32 *posAtrd;
    STTBF grpXstAtnOwners, SttbfAtnbkmk;
    BKF *bkf;
    U32 *posBKF;
    U32 bkf_intervals;
    BKL *bkl;
    U32 *posBKL;
    U32 bkl_intervals;
    wvVersion ver = wvQuerySupported (&ps->fib, NULL);
    external_wvReleasePAPX_FKP ();
    external_wvReleaseCHPX_FKP ();

    /*dop */
    wvGetDOP (ver, &ps->dop, ps->fib.fcDop,
	      ps->fib.lcbDop, ps->tablefd);

#if 0
/* 
this is the versioning name information, the first 22 bytes of each sttbf entry are 
unknown, the rest is a ordinary unicode string, is the time and date and saved by
encoded into the first 22 bytes.
*/
    STTBF versioning;
    if (ver == 0)
      {
	  U16 *str;
	  wvError (("into the versions\n"));
	  wvGetSTTBF (&versioning, ps->fib.fcSttbfUssr, ps->fib.lcbSttbfUssr,
		      ps->tablefd);
	  str = UssrStrBegin (&versioning, 0);
	  wvError (("versioning text is %s\n", wvWideStrToMB (str)));
      }
#endif

    wvGetATRD_PLCF (&atrd, &posAtrd, &atrd_intervals, ps->fib.fcPlcfandRef,
		    ps->fib.lcbPlcfandRef, ps->tablefd);
    wvGetGrpXst (&grpXstAtnOwners, ps->fib.fcGrpXstAtnOwners,
		 ps->fib.lcbGrpXstAtnOwners, ps->tablefd);
    wvTrace (
	     ("offset is %x, len is %d\n", ps->fib.fcSttbfAtnbkmk,
	      ps->fib.lcbSttbfAtnbkmk));
    wvGetSTTBF (&SttbfAtnbkmk, ps->fib.fcSttbfAtnbkmk,
		ps->fib.lcbSttbfAtnbkmk, ps->tablefd);
    wvGetBKF_PLCF (&bkf, &posBKF, &bkf_intervals, ps->fib.fcPlcfAtnbkf,
		   ps->fib.lcbPlcfAtnbkf, ps->tablefd);
    wvGetBKL_PLCF (&bkl, &posBKL, &bkl_intervals, ps->fib.fcPlcfAtnbkl,
           ps->fib.lcbPlcfAtnbkl, ps->fib.fcPlcfAtnbkf, ps->fib.lcbPlcfAtnbkf,
           ps->tablefd);

    /*we will need the stylesheet to do anything useful with layout and look */
    wvGetSTSH (&ps->stsh, ps->fib.fcStshf, ps->fib.lcbStshf, ps->tablefd);

    /* get font list */
    if ((ver == WORD6)
	|| (ver == WORD7))
	wvGetFFN_STTBF6 (&ps->fonts, ps->fib.fcSttbfffn, ps->fib.lcbSttbfffn,
			 ps->tablefd);
    else
	wvGetFFN_STTBF (&ps->fonts, ps->fib.fcSttbfffn, ps->fib.lcbSttbfffn,
			ps->tablefd);

    /*we will need the table of names to answer questions like the name of the doc */
    if ((ver == WORD6)
	|| (ver == WORD7))
      {
	  wvGetSTTBF6 (&ps->anSttbfAssoc, ps->fib.fcSttbfAssoc,
		       ps->fib.lcbSttbfAssoc, ps->tablefd);
	  wvGetSTTBF6 (&ps->Sttbfbkmk, ps->fib.fcSttbfbkmk,
		       ps->fib.lcbSttbfbkmk, ps->tablefd);
      }
    else
      {
	  wvGetSTTBF (&ps->anSttbfAssoc, ps->fib.fcSttbfAssoc,
		      ps->fib.lcbSttbfAssoc, ps->tablefd);
	  wvGetSTTBF (&ps->Sttbfbkmk, ps->fib.fcSttbfbkmk,
		      ps->fib.lcbSttbfbkmk, ps->tablefd);
      }

    /*Extract all the list information that we will need to handle lists later on */
    wvGetLST (&ps->lst, &ps->noofLST, ps->fib.fcPlcfLst, ps->fib.lcbPlcfLst,
	      ps->tablefd);
    wvGetLFO_records (&ps->lfo, &ps->lfolvl, &ps->lvl, &ps->nolfo,
		      &ps->nooflvl, ps->fib.fcPlfLfo, ps->fib.lcbPlfLfo,
		      ps->tablefd);
    /* init the starting list number table */
    if (ps->nolfo)
      {
	  ps->liststartnos = (U32 *) wvMalloc (9 * ps->nolfo * sizeof (U32));
	  ps->listnfcs = (U8 *) wvMalloc (9 * ps->nolfo);
	  ps->finallvl = (LVL *) wvMalloc (9 * ps->nolfo * sizeof (LVL));
	  for (i = 0; i < 9 * ps->nolfo; i++)
	    {
		ps->liststartnos[i] = 0xffffffffL;
		ps->listnfcs[i] = 0xff;
		wvInitLVL (&(ps->finallvl[i]));
	    }
      }
    else
      {
	  ps->liststartnos = NULL;
	  ps->listnfcs = NULL;
	  ps->finallvl = NULL;
      }

    /*Extract Graphic Information */
    wvGetFSPA_PLCF (&ps->fspa, &ps->fspapos, &ps->nooffspa,
		    ps->fib.fcPlcspaMom, ps->fib.lcbPlcspaMom, ps->tablefd);
    wvGetFDOA_PLCF (&ps->fdoa, &ps->fdoapos, &ps->nooffdoa,
		    ps->fib.fcPlcdoaMom, ps->fib.lcbPlcdoaMom, ps->tablefd);

    wvGetCLX (ver, &ps->clx,
	      (U32) ps->fib.fcClx, ps->fib.lcbClx, (U8) ps->fib.fExtChar,
	      ps->tablefd);

    para_fcFirst = char_fcFirst = section_fcFirst =
	wvConvertCPToFC (0, &ps->clx);

#ifdef DEBUG
    if ((ps->fib.ccpFtn) || (ps->fib.ccpHdr))
	wvTrace (("Special ending\n"));
#endif

    /*
       we will need the paragraph and character bounds table to make decisions as 
       to where a table begins and ends
     */
    if ((ver == WORD6)
	|| (ver == WORD7))
      {
	  wvGetBTE_PLCF6 (&btePapx, &posPapx, &para_intervals,
			  ps->fib.fcPlcfbtePapx, ps->fib.lcbPlcfbtePapx,
			  ps->tablefd);
	  wvGetBTE_PLCF6 (&bteChpx, &posChpx, &char_intervals,
			  ps->fib.fcPlcfbteChpx, ps->fib.lcbPlcfbteChpx,
			  ps->tablefd);
      }
    else
      {
	  wvGetBTE_PLCF (&btePapx, &posPapx, &para_intervals,
			 ps->fib.fcPlcfbtePapx, ps->fib.lcbPlcfbtePapx,
			 ps->tablefd);
	  wvGetBTE_PLCF (&bteChpx, &posChpx, &char_intervals,
			 ps->fib.fcPlcfbteChpx, ps->fib.lcbPlcfbteChpx,
			 ps->tablefd);
      }

    wvGetSED_PLCF (&sed, &posSedx, &section_intervals, ps->fib.fcPlcfsed,
		   ps->fib.lcbPlcfsed, ps->tablefd);
    wvTrace (("section_intervals is %d\n", section_intervals));

    wvInitPAPX_FKP (&para_fkp);
    wvInitCHPX_FKP (&char_fkp);

    if(wvHandleDocument (ps, DOCBEGIN))
		goto  finish_processing;

	/*get stream size for bounds checking*/
	stream_size = wvStream_size(ps->mainfd);

    /*for each piece */
    for (piececount = 0; piececount < ps->clx.nopcd; piececount++)
      {
	  ichartype =
	      wvGetPieceBoundsFC (&beginfc, &endfc, &ps->clx, piececount);
	  if(ichartype==-1)
		  break;
	  chartype = (U8) ichartype;
	  /*lvm007@aha.ru fix antiloop: check stream size */
	  if(beginfc>stream_size || endfc>stream_size){
		  wvError (
		   ("Piece Bounds out of range!, its a disaster\n"));
		  continue;
	  }
	  wvStream_goto (ps->mainfd, beginfc);
	  /*lvm007@aha.ru fix antiloop fix*/
	  if(wvGetPieceBoundsCP (&begincp, &endcp, &ps->clx, piececount)==-1)
		  break;
	  wvTrace (
		   ("piece begins at %x and ends just before %x. the char end is %x\n",
		    beginfc, endfc, char_fcLim));

	  /*
	     text that is not in the same piece is not guaranteed to have the same properties as
	     the rest of the exception run, so force a stop and restart of these properties.
	   */
	  char_fcLim = beginfc;

	  for (i = begincp, j = beginfc; (i < endcp /*&& i<ps->fib.ccpText */ );
	       i++, j += wvIncFC (chartype))
	    {
		ps->currentcp = i;
		/* character properties */
		if (j == char_fcLim)
		  {
		      wvHandleElement (ps, CHARPROPEND, (void *) &achp,
				       char_dirty);
		      char_pendingclose = 0;
		  }

		/* comment ending location */
		if (i == comment_cpLim)
		  {
		      wvHandleElement (ps, COMMENTEND, (void *) catrd, 0);
		      comment_pendingclose = 0;
		  }

		/* paragraph properties */
		if (j == para_fcLim)
		  {
		      wvHandleElement (ps, PARAEND, (void *) &apap, para_dirty);
		      para_pendingclose = 0;
		  }

		/* section properties */
		if (j == section_fcLim)
		  {
		      wvHandleElement (ps, SECTIONEND, (void *) &sep,
				       section_dirty);
		      section_pendingclose = 0;
		  }

		if ((section_fcLim == 0xffffffff) || (section_fcLim == j))
		  {
		      section_dirty =
			  wvGetSimpleSectionBounds (ver, ps,
						    &sep, &section_fcFirst,
						    &section_fcLim, i,
						    &ps->clx, sed, &spiece,
						    posSedx,
						    section_intervals,
						    &ps->stsh, ps->mainfd);
		      section_dirty =
			  (wvGetComplexSEP
			   (ver, &sep, spiece,
			    &ps->stsh, &ps->clx) ? 1 : section_dirty);
		  }

		if (j == section_fcFirst)
		  {
		      wvHandleElement (ps, SECTIONBEGIN, (void *) &sep,
				       section_dirty);
		      section_pendingclose = 1;
		  }


		if ((para_fcLim == 0xffffffffL) || (para_fcLim == j))
		  {
		      wvReleasePAPX_FKP (&para_fkp);
		      wvTrace (
			       ("cp and fc are %x(%d) %x\n", i, i,
				wvConvertCPToFC (i, &ps->clx)));
		      cpiece =
			  wvGetComplexParaBounds (ver, &para_fkp,
						  &para_fcFirst, &para_fcLim,
						  wvConvertCPToFC (i,
								   &ps->clx),
						  &ps->clx, btePapx, posPapx,
						  para_intervals, piececount,
						  ps->mainfd);
		      wvTrace (
			       ("para begin and end is %x %x\n", para_fcFirst,
				para_fcLim));

		      if (0 == para_pendingclose)
			{
			    /*
			       if there's no paragraph open, but there should be then I believe that the fcFirst search
			       has failed me, so I set it to now. I need to investigate this further. I believe it occurs
			       when a the last piece ended simultaneously with the last paragraph, and that the algorithm
			       for finding the beginning of a para breaks under that condition. I need more examples to
			       be sure, but it happens is very large complex files so its hard to find
			     */
			    if (j != para_fcFirst)
			      {
				  wvWarning (
					     ("There is no paragraph due to open but one should be, plugging the gap.\n"));
				  para_fcFirst = j;
			      }
			}

		  }

		if (j == para_fcFirst)
		  {
		      para_dirty =
			  wvAssembleSimplePAP (ver, &apap, para_fcLim, &para_fkp, ps);
		      para_dirty =
			  (wvAssembleComplexPAP
			   (ver, &apap, cpiece, ps) ? 1 : para_dirty);
#ifdef SPRMTEST
		      {
			  int p;
			  wvTrace (("Assembled Complex\n"));
			  for (p = 0; p < apap.itbdMac; p++)
			      wvError (
				       ("Tab stop positions are %f inches (%d)\n",
					((float) (apap.rgdxaTab[p])) / 1440,
					apap.rgdxaTab[p]));
		      }
#endif

		      /* test section */
		      wvReleasePAPX_FKP (&para_fkp);
		      wvTrace (
			       ("cp and fc are %x(%d) %x\n", i, i,
				wvConvertCPToFC (i, &ps->clx)));
		      npiece =
			  wvGetComplexParaBounds (ver, &para_fkp,
						  &dummy, &nextpara_fcLim,
						  para_fcLim, &ps->clx,
						  btePapx, posPapx,
						  para_intervals, piececount,
						  ps->mainfd);
		      wvTrace (
			       ("para begin and end is %x %x\n", para_fcFirst,
				para_fcLim));
		      if (npiece > -1)
			{
			    wvAssembleSimplePAP (ver, &ps->nextpap, nextpara_fcLim, &para_fkp, ps);
			    wvAssembleComplexPAP (ver, &ps->nextpap, npiece,ps);
			}
		      else
			  wvInitPAP (&ps->nextpap);
		      /* end test section */

		      if ((apap.fInTable) && (!apap.fTtp))
			{
			    wvGetComplexFullTableInit (ps, para_intervals,
						       btePapx, posPapx,
						       piececount);
			    wvGetComplexRowTap (ps, &apap, para_intervals,
						btePapx, posPapx, piececount);
			}
		      else if (apap.fInTable == 0)
			  ps->intable = 0;

		      wvHandleElement (ps, PARABEGIN, (void *) &apap,
				       para_dirty);

		      char_fcLim = j;
		      para_pendingclose = 1;
		  }


		if ((comment_cpLim == 0xffffffffL) || (comment_cpLim == i))
		  {
		      wvTrace (
			       ("searching for the next comment begin cp is %d\n",
				i));
		      catrd =
			  wvGetCommentBounds (&comment_cpFirst,
					      &comment_cpLim, i, atrd,
					      posAtrd, atrd_intervals,
					      &SttbfAtnbkmk, bkf, posBKF,
					      bkf_intervals, bkl, posBKL,
					      bkl_intervals);
		      wvTrace (
			       ("begin and end are %d %d\n", comment_cpFirst,
				comment_cpLim));
		  }

		if (i == comment_cpFirst)
		  {
		      wvHandleElement (ps, COMMENTBEGIN, (void *) catrd, 0);
		      comment_pendingclose = 1;
		  }


		if ((char_fcLim == 0xffffffffL) || (char_fcLim == j))
		  {
		      wvReleaseCHPX_FKP (&char_fkp);
		      /*try this without using the piece of the end char for anything */
		      wvGetComplexCharBounds (ver, &char_fkp,
					      &char_fcFirst, &char_fcLim,
					      wvConvertCPToFC (i, &ps->clx),
					      &ps->clx, bteChpx, posChpx,
					      char_intervals, piececount,
					      ps->mainfd);
		      wvTrace (
			       ("Bounds from %x to %x\n", char_fcFirst,
				char_fcLim));
		      if (char_fcLim == char_fcFirst)
			  wvError (
				   ("I believe that this is an error, and you might see incorrect character properties\n"));
		      if (0 == char_pendingclose)
			{
			    /*
			       if there's no character run open, but there should be then I believe that the fcFirst search
			       has failed me, so I set it to now. I need to investigate this further.
			     */
			    if (j != char_fcFirst)
			      {
				  wvTrace (
					   ("There is no character run due to open but one should be, plugging the gap.\n"));
				  char_fcFirst = j;
			      }

			}
		      else{
  			 /* lvm007@aha.ru fix: if currentfc>fcFirst but CHARPROP's changed look examples/charprops.doc for decode_simple*/
			 if(char_fcFirst< j)
				char_fcFirst = j;
		       }
		  }

		if (j == char_fcFirst)
		  {
		      /* a CHP's base style is in the para style */
		      /*achp.istd = apap.istd;*/
		      wvTrace (("getting chp\n"));
		      char_dirty =
				  wvAssembleSimpleCHP (ver, &achp, &apap,
					       char_fcLim, &char_fkp,
					       &ps->stsh);
		      wvTrace (("getting complex chp\n"));
		      char_dirty =
			  (wvAssembleComplexCHP
			   (ver, &achp, cpiece,
			    &ps->stsh, &ps->clx) ? 1 : char_dirty);
		      wvHandleElement (ps, CHARPROPBEGIN, (void *) &achp,
				       char_dirty);
		      char_pendingclose = 1;
		  }


		eachchar = wvGetChar (ps->mainfd, chartype);

		/* previously, in place of ps there was a NULL,
		 * but it was crashing Abiword. Was it NULL for a
		 * reason? -JB */
		/* 
		   nah, it was a oversight from when i didn't actually
		   use ps in this function
		   C.
		 */
		if ((eachchar == 0x07) && (!achp.fSpec))
		    ps->endcell = 1;

		wvTrace (("char pos is %x %x\n", j, eachchar));
		wvOutputTextChar (eachchar, chartype, ps, &achp);
	    }

	  if (j == para_fcLim)
	    {
		wvHandleElement (ps, PARAEND, (void *) &apap, para_dirty);
		para_pendingclose = 0;
		para_fcLim = 0xffffffffL;
	    }

	  if (i == comment_cpLim)
	    {
		wvHandleElement (ps, COMMENTEND, (void *) catrd, 0);
		comment_pendingclose = 0;
		comment_cpLim = 0xffffffffL;
	    }

	  if (j == char_fcLim)
	    {
		wvHandleElement (ps, CHARPROPEND, (void *) &achp, char_dirty);
		char_pendingclose = 0;
		char_fcLim = 0xffffffffL;
	    }

#if 0
	  /*      
	     I might have to rethink this closing tag enforcer for complex mode, have to think the
	     flow out a bit more, this section one is plain wrong, im leaving it here so i won't
	     forget and be tempted to put it back in :-)
	     if (j == section_fcLim)
	     {
	     wvHandleElement(ps, SECTIONEND, (void*)&sep,section_dirty);
	     section_pendingclose=0;
	     }
	   */
#endif
      }

 finish_processing:
    if (char_pendingclose)
      {
	  wvInitCHP (&achp);
	  wvHandleElement (ps, CHARPROPEND, (void *) &achp, char_dirty);
      }

    if (comment_pendingclose)
	wvHandleElement (ps, COMMENTEND, (void *) catrd, 0);

    if (para_pendingclose)
      {
	  wvInitPAP (&apap);
	  wvHandleElement (ps, PARAEND, (void *) &apap, para_dirty);
      }

    if (section_pendingclose)
	wvHandleElement (ps, SECTIONEND, (void *) &sep, section_dirty);

    wvFree (ps->fspa);
    wvFree (ps->fspapos);
    wvFree (ps->fdoa);
    wvFree (ps->fdoapos);

    wvFree (posBKL);
    wvFree (bkl);
    wvFree (posBKF);
    wvFree (bkf);
    wvFree (posAtrd);
    wvFree (atrd);

    wvReleasePAPX_FKP (&para_fkp);
    wvReleaseCHPX_FKP (&char_fkp);

    wvHandleDocument (ps, DOCEND);
    wvFree (posSedx);
    wvFree (sed);

    wvFree (ps->liststartnos);
    wvFree (ps->listnfcs);
    for (i = 0; i < 9 * ps->nolfo; i++)
	wvReleaseLVL (&(ps->finallvl[i]));
    wvFree (ps->finallvl);

    wvReleaseLST (&ps->lst, ps->noofLST);
    wvReleaseLFO_records (&ps->lfo, &ps->lfolvl, &ps->lvl, ps->nooflvl);
    wvReleaseSTTBF (&ps->anSttbfAssoc);

    wvFree (btePapx);
    wvFree (posPapx);
    wvFree (bteChpx);
    wvFree (posChpx);
    wvReleaseCLX (&ps->clx);
    wvReleaseFFN_STTBF (&ps->fonts);
    wvReleaseSTSH (&ps->stsh);
    wvReleaseSTTBF (&SttbfAtnbkmk);
    wvReleaseSTTBF (&grpXstAtnOwners);
    if (ps->vmerges)
      {
	  for (i = 0; i < ps->norows; i++)
	      wvFree (ps->vmerges[i]);
	  wvFree (ps->vmerges);
      }
    wvFree (ps->cellbounds);
	wvOLEFree(ps);
    tokenTreeFreeAll ();
}

/*
 The process thus far has created a SEP that describes what the section properties of 
 the section at the last full save. 

 1) Now apply any section sprms that were linked to the piece that contains the 
 section's section mark. 
 
 2) If pcd.prm.fComplex is 0, pcd.prm contains 1 sprm which should be applied to 
 the local SEP if it is a section sprm. 
 
 3) If pcd.prm.fComplex is 1, pcd.prm.igrpprl is the index of a grpprl in the CLX. 
 If that grpprl contains any section sprms, they should be applied to the local SEP
*/
int
wvGetComplexSEP (wvVersion ver, SEP * sep, U32 cpiece, STSH * stsh, CLX * clx)
{
    int ret = 0;
    U16 sprm, pos = 0, i = 0;
    U8 *pointer;
    U16 index;
    U8 val;
    Sprm RetSprm;

    if (clx->pcd[cpiece].prm.fComplex == 0)
      {
	  val = clx->pcd[cpiece].prm.para.var1.val;
	  pointer = &val;
#ifdef SPRMTEST
	  wvError (("singleton\n", clx->pcd[cpiece].prm.para.var1.isprm));
#endif
	  RetSprm =
	      wvApplySprmFromBucket (ver,
				     (U16) wvGetrgsprmPrm ( (U16) clx->pcd[cpiece].prm.
						     para.var1.isprm), NULL,
				     NULL, sep, stsh, pointer, &pos, NULL);
	  if (RetSprm.sgc == sgcSep)
	      ret = 1;
      }
    else
      {
	  index = clx->pcd[cpiece].prm.para.var2.igrpprl;
#ifdef SPRMTEST
	  fprintf (stderr, "\n");
	  while (i < clx->cbGrpprl[index])
	    {
		fprintf (stderr, "%x (%d)\n", *(clx->grpprl[index] + i),
			 *(clx->grpprl[index] + i));
		i++;
	    }
	  fprintf (stderr, "\n");
	  i = 0;
#endif
	  while (i < clx->cbGrpprl[index])
	    {
		if (ver == WORD8)
		    sprm = bread_16ubit (clx->grpprl[index] + i, &i);
		else
		  {
		      sprm = bread_8ubit (clx->grpprl[index] + i, &i);
		      sprm = (U8) wvGetrgsprmWord6 ( (U8) sprm);
		  }
		pointer = clx->grpprl[index] + i;
		RetSprm =
		    wvApplySprmFromBucket (ver, sprm, NULL, NULL, sep, stsh,
					   pointer, &i, NULL);
		if (RetSprm.sgc == sgcSep)
		    ret = 1;
	    }
      }
    return (ret);
}

/*
The process thus far has created a PAP that describes
what the paragraph properties of the paragraph were at the last full save.

1) Now it's necessary to apply any paragraph sprms that were linked to the
piece that contains the paragraph's paragraph mark. 

2) If pcd.prm.fComplex is 0, pcd.prm contains 1 sprm which should only be 
applied to the local PAP if it is a paragraph sprm. 

3) If pcd.prm.fComplex is 1, pcd.prm.igrpprl is the index of a grpprl in the 
CLX.  If that grpprl contains any paragraph sprms, they should be applied to 
the local PAP.
*/
int
wvAssembleComplexPAP (wvVersion ver, PAP * apap, U32 cpiece, wvParseStruct *ps)
{
    int ret = 0;
    U16 sprm, pos = 0, i = 0;
    U8 sprm8;
    U8 *pointer;
    U16 index;
    U8 val;
    Sprm RetSprm;

    if (ps->clx.pcd[cpiece].prm.fComplex == 0)
      {
	  val = ps->clx.pcd[cpiece].prm.para.var1.val;
	  pointer = &val;
#ifdef SPRMTEST
	  wvError (("singleton\n", ps->clx.pcd[cpiece].prm.para.var1.isprm));
#endif
	  RetSprm =
	      wvApplySprmFromBucket (ver,
				     (U16) wvGetrgsprmPrm ( (U16) ps->clx.pcd[cpiece].prm.
						     para.var1.isprm), apap,
				     NULL, NULL, &ps->stsh, pointer, &pos, ps->data);
	  if (RetSprm.sgc == sgcPara)
	      ret = 1;
      }
    else
      {
	  index = ps->clx.pcd[cpiece].prm.para.var2.igrpprl;
#ifdef SPRMTEST
	  wvError (("HERE-->\n"));
	  fprintf (stderr, "\n");
	  for (i = 0; i < ps->clx.cbGrpprl[index]; i++)
	      fprintf (stderr, "%x ", *(ps->clx.grpprl[index] + i));
	  fprintf (stderr, "\n");
	  i = 0;
#endif
	  while (i < ps->clx.cbGrpprl[index])
	    {
		if (ver == WORD8)
		    sprm = bread_16ubit (ps->clx.grpprl[index] + i, &i);
		else
		  {
		      sprm8 = bread_8ubit (ps->clx.grpprl[index] + i, &i);
		      sprm = (U16) wvGetrgsprmWord6 (sprm8);
		      wvTrace (("sprm is %x\n", sprm));
		  }
		pointer = ps->clx.grpprl[index] + i;
		RetSprm =
		    wvApplySprmFromBucket (ver, sprm, apap, NULL, NULL, &ps->stsh,
					   pointer, &i, ps->data);
		if (RetSprm.sgc == sgcPara)
		    ret = 1;
	    }
      }
    return (ret);
}

/* CHP version of the above. follows the same rules -JB */
int
wvAssembleComplexCHP (wvVersion ver, CHP * achp, U32 cpiece, STSH * stsh,
		      CLX * clx)
{
    int ret = 0;
    U16 sprm, pos = 0, i = 0;
    U8 sprm8;
    U8 *pointer;
    U16 index;
    U8 val;
    Sprm RetSprm;

    if (clx->pcd[cpiece].prm.fComplex == 0)
      {
	  val = clx->pcd[cpiece].prm.para.var1.val;
	  pointer = &val;
#ifdef SPRMTEST
	  wvError (("singleton %d\n", clx->pcd[cpiece].prm.para.var1.isprm));
#endif
	  RetSprm =
	      wvApplySprmFromBucket (ver,
				     (U16) wvGetrgsprmPrm ( (U16) clx->pcd[cpiece].prm.
						     para.var1.isprm), NULL,
				     achp, NULL, stsh, pointer, &pos, NULL);
	  if (RetSprm.sgc == sgcChp)
	      ret = 1;
      }
    else
      {
	  index = clx->pcd[cpiece].prm.para.var2.igrpprl;
#ifdef SPRMTEST
	  fprintf (stderr, "\n");
	  for (i = 0; i < clx->cbGrpprl[index]; i++)
	      fprintf (stderr, "%x ", *(clx->grpprl[index] + i));
	  fprintf (stderr, "\n");
	  i = 0;
#endif
	  while (i < clx->cbGrpprl[index])
	    {
		if (ver == WORD8)
		    sprm = bread_16ubit (clx->grpprl[index] + i, &i);
		else
		  {
		      sprm8 = bread_8ubit (clx->grpprl[index] + i, &i);
		      sprm = (U16) wvGetrgsprmWord6 (sprm8);
		  }
		pointer = clx->grpprl[index] + i;
		RetSprm =
		    wvApplySprmFromBucket (ver, sprm, NULL, achp, NULL, stsh,
					   pointer, &i, NULL);
		if (RetSprm.sgc == sgcChp)
		    ret = 1;
	    }
      }
    return (ret);
}

```
