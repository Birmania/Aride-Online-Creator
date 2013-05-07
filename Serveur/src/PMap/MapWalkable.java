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

import Communications.OutputBuffer;
import Enumerations.Directions;

public abstract class MapWalkable extends MapMovable {
	
	
	public MapWalkable(Map.MapInstance map, byte x, byte y)
	{
		super(map, x, y, Directions.DIR_DOWN);
	}
	
	public abstract OutputBuffer getStartMovePacket();
	
	protected abstract OutputBuffer getStopMovePacket();
	
	public abstract void sendStopMovePacketToMap();
	
	public abstract OutputBuffer getDirMovePacket();
	
	public abstract OutputBuffer getDirPacket();
	
	public byte moveOneOnDir()
	{
		byte rval = 0;
		byte x = this.getX();
		byte y = this.getY();
		
		
		switch (this.getDir())
		{
			case DIR_UP:
				y--;
				break;
			case DIR_DOWN:
				y++;
				break;
			case DIR_LEFT:
				x--;
				break;
			case DIR_RIGHT:
				x++;
				break;
		}
		
		if (x == -1 || y == -1 || x == this.getMapInstance().getMap().mapAttributes.tiles.length || y == this.getMapInstance().getMap().mapAttributes.tiles[0].length)
		{
			//TODO : Look if a problem can occur if we delete this two lines
			//this.stopMovementTimer();
			//this.sendStopMovePacketToMap();
			rval = 1;
		}
		else
		{
			if (this.getMapInstance().getTileAllocation(x, y).isTraversableBy(this))
			{
				this.setXY(x, y);
				
				// Tester si il n'y a pas un projectile
				// TODO : Locker le joueur lors d'une perte de vie car on pourrait très bien parcourir et lui faire perdre de la vie et avoir en même temps de nouvelle projectiles qui viennent de l'extérieur
				
				this.getMapInstance().checkConflict(x, y);
			}
			else
			{
				this.stopMovementTimer();
				this.sendStopMovePacketToMap();
				rval = 1;
			}
		}
		return rval;
	}
}
