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

import Main.World;

public enum ItemTypes {
	ItemTypeWeapon(1), ItemTypeThrowable(2), ItemTypeMissile(3), ItemTypeArmor(4), ItemTypeHelmet(5), ItemTypeShield(6), ItemTypePotion(7);
	 
	 private byte code;
	 
	 private ItemTypes(int c) {
	   code = (byte)c;
	 }
	 
	 public byte getCode() {
	   return code;
	 }
	 
	 static public boolean isWeapon(short itemId)
	 {
		 boolean rval= false;
		 
		 if (World.getInstance().items.get(itemId).type >= ItemTypes.ItemTypeWeapon.getCode() && World.getInstance().items.get(itemId).type <= ItemTypes.ItemTypeMissile.getCode())
		 {
			 rval = true;
		 }
		 
		 return rval;
	 }
	 
	 static public boolean isArmor(short itemId)
	 {
		 return World.getInstance().items.get(itemId).type == ItemTypes.ItemTypeArmor.getCode();
	 }
	 
	 static public boolean isHelmet(short itemId)
	 {
		 return World.getInstance().items.get(itemId).type == ItemTypes.ItemTypeHelmet.getCode();
	 }
	 
	 static public boolean isShield(short itemId)
	 {
		 return World.getInstance().items.get(itemId).type == ItemTypes.ItemTypeShield.getCode();
	 }
	 
	 static public boolean isPotion(short itemId)
	 {
		 return World.getInstance().items.get(itemId).type == ItemTypes.ItemTypePotion.getCode();
	 }
}
