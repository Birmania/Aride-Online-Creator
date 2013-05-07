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

package Main;

import Communications.BinaryFile;
import Communications.InputBuffer;

public class Item extends BinaryFile {
	public String name;
	public String desc;
	public int pic;
	public byte type;
	public int datas[];
	
	public short strReq;
	public short defReq;
	public short dexReq;
	public short sciReq;
	public short langReq;
	
	public byte empilable;
	
	public short lifeEffect;
	public short sleepEffect;
	public short staminaEffect;
	public short addHP;
	public short addSLP;
	public short addSTP;
	public short addStr;
	public short addDef;
	public short addSci;
	public short addDex;
	public short addLang;
	public int addExp;
	
	// TODO : Le mettre dans les datas
	public int attackSpeed;
	
	public int color;
	
	public byte sex;
	
	public Item(InputBuffer buffer)
	{
		this.deserialize(buffer);
	}
}
