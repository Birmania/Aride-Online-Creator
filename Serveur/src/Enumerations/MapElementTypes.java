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

import Main.Player;
import PMap.MapElement;

public enum MapElementTypes {
	 Player(0), Npc(1), Pet(2);
	 
	 private byte code;
	 
	 private MapElementTypes(int c) {
	   code = (byte)c;
	 }
	 
	 public byte getCode() {
	   return code;
	 }
	 
	 public static MapElementTypes get(MapElement element)
	 {
		 MapElementTypes rval = null;
		 
		 if (element instanceof Player)
		 {
			 rval = Player;
		 }
		 else if (element instanceof Npc.Npc)
		 {
			 rval = Npc;
		 }
		 else if (element instanceof Main.Pet)
		 {
			 rval = Pet;
		 }
		 
		 return rval;
	 }
}