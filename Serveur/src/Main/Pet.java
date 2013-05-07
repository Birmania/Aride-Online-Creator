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

import java.util.Observable;
import java.util.Observer;

import Communications.OutputBuffer;
import Communications.Transmission;
import Exceptions.NoPlayerException;
import Interfaces.IKillable;
import Interfaces.IRecognizable;
import Miscs.MessageLogger;
import Miscs.TalkingRunnable;
import Npc.NpcType;
import PMap.Map;
import PMap.MapFighter;

public class Pet extends MapFighter implements Observer, IRecognizable{

	public short localId;
	
	public NpcType type;	// Pas obligatoire
	
	public boolean isFollowing;
	
	public Pet(short localId, Map.MapInstance map, byte x, byte y, NpcType type) {
		super(map, x, y, type.maxHp, type.attackSpeed); // life then attackspeed
		this.localId = localId;
		this.type = type;
		this.isFollowing = true;
	}
	
	public void writeStartInfosInPacket(OutputBuffer packet)
	{	
		packet.writeShort(this.type.id);
	}
	
	// TODO : Faire comme pour le client, un writeInPacket et un WriteStartInfosInPacket
	public void writePositionInPacket(OutputBuffer packet)
	{
		packet.writeByte(this.getX());
		packet.writeByte(this.getY());
		packet.writeByte(this.getDir().getCode());
		if (this.moving)
		{
			packet.writeByte(this.speed);
		}
		else
		{
			packet.writeByte((byte)0);
		}
	}

	@Override
	public short getId() {
		return this.localId;
	}

	@Override
	public void update(Observable arg0, Object arg1) {
		if (this.tryToPrepareToFight())
		{
			if (!this.isDead())
			{
				this.attackTarget();
			}
			this.escapeFromFight();
		}
		
	}

	@Override
	public String getName() {
		return this.type.name;
	}

	@Override
	public void clearAll() {
		Runnable remover = new TalkingRunnable(new Runnable() {
			
			public void run() {
				try {
					Population.getInstance().playersLock.readLock().lock();
					Player master = Population.getInstance().getPlayer(Pet.this.getId()).player;

					master.petLock.lock();
					
					Pet.this.getMapInstance().deleteObserver(Pet.this);
					if (Pet.this.moving)
					{
						Pet.this.stopMovementTimer();
					}

					Pet.this.stopAttackTimer();
					Pet.this.getMapInstance().removePet(Pet.this);
					
					master.pet = null;
					
					Transmission.sendToAllPlayer(Pet.this.getDeadPacket());
					master.petLock.unlock();

					Population.getInstance().playersLock.readLock().unlock();
				} catch (NoPlayerException e) {
					MessageLogger.getInstance().log(e);
				}
			}
		});
		ServerConfiguration.getInstance().scheduledExecutor.submit(remover);
	}

	//@Override
	public OutputBuffer getDeadPacket() {
		OutputBuffer packet = new OutputBuffer("SPetDead");
		packet.writeShort(this.localId);
		return packet;
	}

	@Override
	public int getDamage() {
		// TODO Auto-generated method stub
		return 5;
	}
	
	@Override
	public void attack(IKillable target) {
		if (!this.isFollowing)
		{
			super.attack(target);
		}
		else
		{
			this.stopAttackTimer();
		}
	}

	@Override
	public OutputBuffer getStartMovePacket() {
		OutputBuffer packet = new OutputBuffer("SPetStartMove");
		
		packet.writeShort(this.localId);
		
		packet.writeByte(this.getDir().getCode());

		packet.writeByte(this.speed);
		
		return packet;
	}

	@Override
	protected OutputBuffer getStopMovePacket() {
		OutputBuffer packet = new OutputBuffer("SPetStopMove");
		
		packet.writeShort(this.localId);
		
		packet.writeByte(this.getX());

		packet.writeByte(this.getY());
		
		return packet;
	}

	@Override
	public OutputBuffer getDirMovePacket() {
		OutputBuffer packet = new OutputBuffer("SPetDirMove");
		
		packet.writeShort(this.localId);
		
		packet.writeByte(this.getDir().getCode());
		 
		packet.writeByte(this.getX()); // X position
		packet.writeByte(this.getY()); // Y position
		
		return packet;
	}

	@Override
	public void sendStopMovePacketToMap() {
		Transmission.sendToMapInstance(this.getMapInstance(), this.getStopMovePacket());
	}
	
	@Override
	public boolean isDead()
	{
		boolean rval = super.isDead();

		try {
			Population.getInstance().getPlayer(this.localId);
		} catch (NoPlayerException e) {
			rval = true;
		}
		
		return rval;
	}
	
	public void followMaster() {
		this.prepareToFight();

		if (!this.isDead())
		{
			try {

				this.isFollowing = true;
				
				this.target = Population.getInstance().getPlayer(this.getId()).player;

				
				this.attackTarget();

			} catch (NoPlayerException e) {
				MessageLogger.getInstance().log(e);
			}
		}
		this.escapeFromFight();
	}

	@Override
	public int getExpValue() {
		return 0;
	}

	@Override
	public OutputBuffer getDirPacket() {
		OutputBuffer packet = new OutputBuffer("SPetDir");
		
		packet.writeShort(this.getId());
		
		packet.writeByte(this.getDir().getCode());
		
		return packet;
	}

}
