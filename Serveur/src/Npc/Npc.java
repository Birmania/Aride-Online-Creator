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

package Npc;

import java.util.Iterator;
import java.util.Observable;
import java.util.Observer;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.ScheduledFuture;
import java.util.concurrent.TimeUnit;

import Communications.OutputBuffer;
import Communications.Transmission;
import Enumerations.Directions;
import Enumerations.ItemTypes;
import Enumerations.NpcBehaviors;
import Interfaces.IKillable;
import Main.Faun;
import Main.Item;
import Main.ItemSlot;
import Main.Player;
import Main.Position;
import Main.ServerConfiguration;
import Main.World;
import Miscs.TalkingRunnable;
import PMap.Map;
import PMap.MapFighter;
import PMap.MapItem;

public class Npc extends MapFighter implements Observer{
	
	public short localId;
	
	public NpcType type;	// Pas obligatoire
	
	protected ScheduledExecutorService randomMovementTimer;
	private Runnable randomMovement;
	private ScheduledFuture<?> randomMovementScheduler;
	
	public Npc(short localId, Map.MapInstance map, byte x, byte y, NpcType type)
	{
		super(map, x, y, type.maxHp, type.attackSpeed);
		
		this.localId = localId;
		
		this.type = type;
		
		this.randomMovementTimer = Executors.newSingleThreadScheduledExecutor();
		this.randomMovement = new TalkingRunnable(new Runnable() {

			@Override
			public void run() {
				Npc.this.testRandomMove();
			}
			
		});
		
		this.launchRandomMovementTimer();
	}
	
	public void launchRandomMovementTimer()
	{
		//this.randomMovementTimer = Executors.newSingleThreadScheduledExecutor();
		this.randomMovementScheduler = this.randomMovementTimer.scheduleAtFixedRate(this.randomMovement, 5000, 5000, TimeUnit.MILLISECONDS);
	}
	
	public void stopRandomMovementTimer()
	{
		//this.randomMovementTimer.shutdownNow();
		this.randomMovementScheduler.cancel(false);
	}

	protected void testRandomMove()
	{	
		this.getMapInstance().lock.lock();
		Npc.this.prepareToFight(); // To not be killed during a test
		if (!Npc.this.isDead())
		{
			if (this.type.range > 0 && !this.attacking && !this.moving)
			{	
				// Si une cible se déconnecte, pas d'update du NPC donc il ne peut pas "perdre de vue" le joueur. On doit le forcer ici
				this.target = null;
				
				byte xMin = (byte)(this.getX() - this.type.range);
				byte yMin =  (byte)(this.getY() - this.type.range);
				byte xMax = (byte)(this.getX() + this.type.range);
				byte yMax = (byte)(this.getY() + this.type.range);
				
				if (xMin < 0)
				{
					xMin = 0;
				}
				if (yMin < 0)
				{
					yMin = 0;
				}

				if (xMax > this.getMapInstance().getMaxX())
				{
					xMax = this.getMapInstance().getMaxX();
				}
				if (yMax > this.getMapInstance().getMaxY())
				{
					yMax = this.getMapInstance().getMaxY();
				}
				
				byte x;
				byte y;
				do
				{
					x = (byte)((Math.random()*(xMax+1-xMin))+xMin);
					y = (byte)((Math.random()*(yMax+1-yMin))+yMin);
				} while (this.getMapInstance().getTileAllocation(x, y).size() > 0);
				
				this.destination = new Position(x, y);
				
				Directions direction = this.getTargetDirection(this.destination);
				if (direction != null)
				{
					this.setDir(direction);

					this.speed = 4;
					this.launchMovementTimer();

					Transmission.sendToMapInstance(this.getMapInstance(), this.getStartMovePacket());
				}
			}
		}
		Npc.this.escapeFromFight();
		this.getMapInstance().lock.unlock();
	}
		
	@Override
	public int getDamage() {
		return 2;
	}

	@Override
	public String getName() {
		return Faun.getInstance().faun[this.localId].name;
	}
	
	public void writeInPacket(OutputBuffer packet)
	{
		packet.writeShort(this.localId);
		packet.writeShort(this.type.id);
		packet.writeByte(this.getX());
		packet.writeByte(this.getY());
		packet.writeInt(this.life);
		
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
	public OutputBuffer getStopMovePacket() {
		// Send the information to the players
		OutputBuffer packet = new OutputBuffer("SNpcStopMove");
		
		packet.writeShort(this.localId);
		
		packet.writeByte(this.getX());

		packet.writeByte(this.getY());
		
		return packet;
	}
	
	@Override
	public OutputBuffer getStartMovePacket() {
		OutputBuffer packet = new OutputBuffer("SNpcStartMove");
		
		packet.writeShort(this.localId);
		
		packet.writeByte(this.getDir().getCode());

		//packet.writeByte((byte)(1000/this.movementTimer.getDelay()));
		packet.writeByte(this.speed);
		
		return packet;
	}
	
	@Override
	public OutputBuffer getDirMovePacket() {
		OutputBuffer packet = new OutputBuffer("SNpcDirMove");
		
		packet.writeShort(this.localId);
		
		packet.writeByte(this.getDir().getCode());
		 
		packet.writeByte(this.getX()); // X position
		packet.writeByte(this.getY()); // Y position
		
		return packet;
	}
	
	public void selectTarget()
	{
		this.target = null;
		int distance = this.type.range; // On intialise la distance à la portée

		switch(NpcBehaviors.values()[this.type.behavior])
		{
			case Violent : // Chasse toujours les joueurs
				synchronized(this.getMapInstance().getPlayers())
				{
					// Trouver le joueur le plus proche
					Iterator<Player> itePla = this.getMapInstance().getPlayers().iterator();
					
					IKillable choosenTarget = null;
					while (itePla.hasNext())
					{
						Player currentPlayer = itePla.next();
						int difference = this.getPosition().distance(currentPlayer.getPosition());
						if (difference <= distance)
						{
							distance = difference;
							choosenTarget = currentPlayer;
						}
						
						currentPlayer.petLock.lock();
						if (currentPlayer.pet != null)
						{
							difference = this.getPosition().distance(currentPlayer.pet.getPosition());
							if (difference <= distance)
							{
								distance = difference;
								choosenTarget = currentPlayer.pet;
							}
						}
						currentPlayer.petLock.unlock();
					}
					if (choosenTarget != null)
					{
						this.destination = null;
					}
					this.target = choosenTarget;
					
					// La target du npc est initialisé au joueur le plus proche visible
				}
				break;
		}
	}
	
	/*@Override
	public void attack(IKillable target) {
		//if (this.target != null && this.getMapInstance().getPlayers().contains(this.target) && this.getPosition().distance(target.getPosition()) <= 1)
		if (this.target != null && this.getMapInstance() == this.target.getMapInstance() && this.getPosition().distance(target.getPosition()) <= 1)
		{
			super.attack(target);
		}
		else
		{
			this.attackTimer.stop();
		}
	}*/

	@Override
	public void update(Observable arg0, Object arg1) { // Syncrhonisé avec move
		if (this.tryToPrepareToFight())
		{
			if (!this.isDead())
			{
				this.selectTarget();
				this.attackTarget();
			}
			this.escapeFromFight();
		}
	}
	
	public void destroy()
	{
		this.getMapInstance().deleteObserver(this);
		this.stopRandomMovementTimer();
		if (this.moving)
		{
			this.stopMovementTimer();
		}
		this.stopAttackTimer();
		this.getMapInstance().removeNpc(this);
		
		this.getMapInstance().scheduleRespawn(this.localId, this.type.spawnSecs);
		
		Transmission.sendToMapInstance(this.getMapInstance(), this.getDeadPacket());
		// TODO : A supprimer
		//System.gc();
	}

	public void clearAll()
	{
		Runnable remover = new TalkingRunnable(new Runnable() {
			
			public void run() {
				if (Npc.this.type.items != null)
				{
					for (NpcItem current : Npc.this.type.items)
					{
						int luck = (int)(Math.random()*(current.luck))+1;
						
						if (luck == 1)
						{
							short dur = 1;
							Item item = World.getInstance().items.get(current.itemNum);
							if (item.type >= ItemTypes.ItemTypeWeapon.getCode() && item.type <= ItemTypes.ItemTypeShield.getCode())
							{
								dur = (short)item.datas[0];
							}
							
							Npc.this.getMapInstance().addItem(new MapItem(Npc.this.getPosition(), new ItemSlot((short)current.itemNum, (short)current.itemVal, dur)));
						}
					}
				}
				
				Npc.this.destroy();
			}
		});
		ServerConfiguration.getInstance().scheduledExecutor.submit(remover);
	}
	
	//@Override
	public OutputBuffer getDeadPacket() {
		OutputBuffer packet = new OutputBuffer("SNpcDead");
		packet.writeShort(this.localId);
		return packet;
	}

	@Override
	public short getId() {
		return this.localId;
	}

	@Override
	public void sendStopMovePacketToMap() {
		Transmission.sendToMapInstance(this.getMapInstance(), this.getStopMovePacket());
	}

	@Override
	public int getExpValue() {
		return this.type.experience;
	}

	@Override
	public OutputBuffer getDirPacket() {
		OutputBuffer packet = new OutputBuffer("SNpcDir");
		
		packet.writeShort(this.getId());
		
		packet.writeByte(this.getDir().getCode());
		
		return packet;
	}

	/*@Override
	public boolean mustGiveUp() {
		boolean rval = false;
		
		if (this.target.getPosition().distance(this.getPosition()) > this.type.range)
		{
			rval = true;
		}
		
		return rval;
	}*/
}
