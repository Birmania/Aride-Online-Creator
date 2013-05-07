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

public enum WeatherTypes {
	 None(0), Raining(1), Snowing(2), Thundering(3), SandStorming(4);
	 
	 private byte code;
	 
	 private WeatherTypes(int c) {
	   code = (byte)c;
	 }
	 
	 public byte getCode() {
	   return code;
	 }
}
