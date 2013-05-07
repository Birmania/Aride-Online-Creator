/*	Copyright (C) 2013  BRULTET Antoine
	
	This file is part of Aride Online Creator.

    Aride Online Creator is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    Aride Online Creator is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with Aride Online Creator.  If not, see <http://www.gnu.org/licenses/>.
*/

package Enumerations;

public enum Colors {
 Black(0), Blue(1), Green(2), Cyan(3), Red(4), Magenta(5), Brown(6), Grey(7), DarkGrey(8), BrightBlue(9), BrightGreen(10), BrightCyan(11), BrightRed(12), Pink(13), Yellow(14), White(15);
 
 private short code;
 
 private Colors(int c) {
   code = (short)c;
 }
 
 public short getCode() {
   return code;
 }
}