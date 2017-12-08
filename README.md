# PoulinHaddad
Predicting tissue distribution of highly lipophilic compounds.

Poulin, P.; Haddad, S. Advancing Prediction of Tissue Distribution and Volume of Distribution of Highly Lipophilic Compounds from a Simpliﬁed Tissue-Composition-Based Model as a Mechanistic Animal Alternative Method. _J. Pharm. Sci._ __2012__, 101, 2250–2261. 

## Prerequisites
* Microsoft Windows 7 or later
* Microsoft Office 2010 or later
* Microsoft .NET 4+

## Download
Download the latest build from [releases](http://www.github.com/HSL/PoulinHaddad/releases/).

Two XLL files are present in the zip archive for a release. The build (32-bit or 64-bit) of Microsoft Office installed will detemine which of the XLL files should be installed.

## Installation
To install the XLL, refer to _Add or remove an Excel add-in_ on this [Microsoft Office support page](https://support.office.com/en-us/article/Add-or-remove-add-ins-0af570c4-5cf3-4fa9-9b88-403625a0b460).

## Getting Started
Refer to Poulin & Haddad to determine whether the model implemented in this add-in is applicable to your problem domain.

### V<sub>ss</sub>

In a worksheet cell, use the __Vss__ function. The three arguments required are:

1. Log octanol:water partition coefficient, log P<sub>ow</sub>
2. Ionization class (N, WB, SB, A, or Z)
3. Does drug chemical structure consist of at least one oxygen atom? (TRUE or FALSE)

If the drug or chemical is not neutral (not class N), supply a fourth argument: 

4. pK<sub>a</sub>

If the drug or chemical is a zwitterion (class Z), supply a fifth argument: 

5. pK<sub>a</sub>,base

Example:

```
=Vss(7.23, "N", TRUE)
```

### Other Computations

To view other functions available with this add-in, in the Excel function wizard, select __Poulin Haddad__ from the _Category_ drop-down.
