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

package PMap;

import Main.Position;

public class MapElement {
	
	private Position position;
	private Map.MapInstance map;
	
	public MapElement (Map.MapInstance map, byte x, byte y) {
		this.setMapInstance(map);
		this.position = new Position(x, y); // -1 pour ne pas désallouer une case lors du premier setX, setY
	}
	
	public void setMapInstance(Map.MapInstance map)
	{
		// TODO : When call this method, lock the map instance
		this.map = map;
	}
	
	public Map.MapInstance getMapInstance()
	{
		return this.map;
	}
	
	public Position getPosition()
	{
		return this.position;
	}
	
	public byte getX()
	{
		return this.position.getX();
	}
	
	public byte getY()
	{
		return this.position.getY();
	}
	
	public void setX(byte x)
	{	
		this.map.removeTileAllocation(this.getPosition(), this);
		this.position.setX(x);
		this.map.addTileAllocation(this.getPosition(), this);
	}
	
	public void setY(byte y)
	{
		this.map.removeTileAllocation(this.getPosition(), this);
		this.position.setY(y);
		this.map.addTileAllocation(this.getPosition(), this);
	}
	
	public void setXY(Position position)
	{
		this.setXY(position.getX(), position.getY());
	}
	
	public void setXY(byte x, byte y)
	{
		this.map.removeTileAllocation(this.getPosition(), this);
		this.position.setX(x);
		this.position.setY(y);
		this.map.addTileAllocation(this.getPosition(), this);
	}
}
