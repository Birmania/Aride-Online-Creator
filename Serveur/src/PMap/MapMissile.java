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

import java.util.concurrent.locks.Lock;
import java.util.concurrent.locks.ReentrantLock;

import Communications.OutputBuffer;
import Communications.Transmission;
import Enumerations.Directions;
import Interfaces.IFighter;
import Interfaces.IKillable;
import Main.Player;

public class MapMissile extends MapMovable implements IFighter {

	private byte id;
	private byte missileType;
	private Player player;
	private int damage;
	private Lock movementLock;
	
	public MapMissile(PMap.Map.MapInstance map, byte x, byte y, Directions dir, byte missileType, Player player, int damage)
	{
		super(map, x, y, dir);
		
		this.movementLock = new ReentrantLock();
		
		synchronized(this.getMapInstance().missiles)
		{
			this.id = -1;
			byte i = 0;
			while ((this.id == -1) && (i < this.getMapInstance().missiles.size()+1)) // deuxième condition non obligatoire
			{
				if (!this.getMapInstance().missiles.containsKey(i))
				{
					// TODO : Côté client il faut faire un redim du tableau d'arrow pour pas avoir de sortie de tableau si trop de fleche en meme temps
					this.id = i;
					this.getMapInstance().missiles.put(i, this);
				}
				i++;
			}
		}
		
		this.missileType = missileType;
		this.player = player;
		this.damage = damage;
		
		/*this.movementTimer.setInitialDelay(90);
		this.movementTimer.setDelay(90);
		this.movementTimer.start();*/
		this.speed = 11;
		this.launchMovementTimer();
		
		Transmission.sendToMapInstance(map, this.getAppearPacket());
	}
	
	public OutputBuffer getAppearPacket()
	{
		OutputBuffer packet = new OutputBuffer("SMissileAppear");
		
		packet.writeShort(this.player.getId());
		packet.writeByte(this.id);
		packet.writeByte(this.missileType);
		packet.writeByte(this.getDir().getCode());
		packet.writeByte(this.getX());
		packet.writeByte(this.getY());
		
		return packet;
	}
	
	public OutputBuffer getDisappearPacket()
	{
		OutputBuffer packet = new OutputBuffer("SMissileDisappear");
		
		packet.writeByte(this.id);
		
		return packet;
	}

	public void destroyMissile()
	{
		this.movementLock.lock();
		synchronized(this.getMapInstance().missiles)
		{
			this.stopMovementTimer();
			this.getMapInstance().removeTileAllocation(this.getPosition(), this);
			Transmission.sendToMapInstance(this.getMapInstance(), this.getDisappearPacket());
			this.getMapInstance().missiles.remove(this.id);
		}
		this.movementLock.unlock();
	}
	
	@Override
	public void move() {
		this.getMapInstance().lock.lock();
		if (this.movementLock.tryLock())
		{
			//CHANGED this.getMapInstance().lock.lock();
			
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
				this.destroyMissile();
			}
			else
			{
				this.setXY(x, y);
				this.getMapInstance().checkConflict(x, y);
			}
			
			//CHANGED this.getMapInstance().lock.unlock();
			this.movementLock.unlock();
		}
		this.getMapInstance().lock.unlock();
	}

	@Override
	public int getDamage() {
		return this.damage;
	}

	@Override
	synchronized public void attack(IKillable target) {
		this.movementLock.lock(); // Just because mapmissile could be destroy by moving
		if (this.moving)
		{
			this.player.attack(target, MapMissile.this.damage);
			this.destroyMissile();
		}
		this.movementLock.unlock();
		/*final IKillable finalTarget = target;
		new Thread() // Create a thread because attack could be call from a moving method (lock of map instance)
        {
            {
                this.setDaemon(true);
                this.start();
            }
 
            public void run()
            {
				MapMissile.this.player.attack(finalTarget, MapMissile.this.damage);
				MapMissile.this.destroyMissile();
            }
        };*/
	}

	public Player getLauncher() {
		return this.player;
	}

	@Override
	public short getId() {
		return this.id;
	}
	
	public void lockMovement()
	{
		this.movementLock.lock();
	}
	
	public void unlockMovement()
	{
		this.movementLock.unlock();
	}

	/*@Override
	public int compareTo(Object arg0) {
		if (this.id < ((MapMissile)arg0).id)
		{
			return -1;
		}
		else if (this.id > ((MapMissile)arg0).id)
		{
			return 1;
		}
		else
		{
			return 0;
		}
	}*/

}
