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

import java.util.HashMap;

import Communications.BinaryFile;
import Communications.InputBuffer;
import PMap.Map;

public class Dream extends BinaryFile {
	public String name;
	public short beginningMap;
	public short beginningX;
	public short beginningY;
	public short maps[];
	
	
	public Dream(InputBuffer buffer)
	{
		this.deserialize(buffer);
	}
	
	public DreamInstance newInstance()
	{
		return new DreamInstance();
	}
	
	public class DreamInstance
	{
		public HashMap<Short, Map.MapInstance> mapsInstances;
		
		public DreamInstance()
		{
			this.mapsInstances = new HashMap<Short, Map.MapInstance>();
			
			for (short currentMap : Dream.this.maps)
			{
				this.mapsInstances.put(currentMap, World.getInstance().getMap(currentMap).newInstance());
			}
		}
		
		// TODO : devra être supprimé dans la version finale
		@Override
		protected void finalize() throws Throwable {
			super.finalize();
		}
	}
}
