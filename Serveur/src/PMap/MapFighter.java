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

import java.util.ArrayList;
import java.util.Map;
import java.util.TreeMap;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.ScheduledFuture;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.locks.Lock;
import java.util.concurrent.locks.ReentrantLock;

import Communications.OutputBuffer;
import Communications.Transmission;
import Enumerations.Directions;
import Interfaces.IFighter;
import Interfaces.IKillable;
import Main.Player;
import Main.Position;
import Main.ServerConfiguration;
import Miscs.TalkingRunnable;

public abstract class MapFighter extends MapWalkable implements IFighter, IKillable {
	
	public IKillable target;
	
	public int life;
	
	public Lock fightLock;
	
	//public long attackTimer;
	//public Timer attackTimer;
	public ScheduledExecutorService attackTimer;
	private ScheduledFuture<?> attackScheduler;
	private Runnable attack;
	public boolean attacking;
	
	public short attackSpeed;
	
	public Position destination; // Destination lors d'un déplacement sans target
	
	//public Lock lock;
	
	public MapFighter(PMap.Map.MapInstance map, byte x, byte y, int life, short attackSpeed) {
		super(map, x, y);
		this.life = life;
		this.fightLock = new ReentrantLock();
		this.attackSpeed = attackSpeed;
		
		//this.lock = new ReentrantLock();
		
		this.attackTimer = Executors.newSingleThreadScheduledExecutor();
		this.attack = new TalkingRunnable(new Runnable() {

			@Override
			public void run() {
				MapFighter.this.attack(MapFighter.this.target);
			}
			
		});
		this.attacking = false;

		/*this.attackTimer = new Timer(attackSpeed, new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				MapFighter.this.attack(MapFighter.this.target);
			}
		});*/
	}

	public void launchMovementTimer()
	{
		super.launchMovementTimer();
	}
	
	public void stopMovementTimer()
	{
		super.stopMovementTimer();
	}
	
	public void launchAttackTimer()
	{
		this.attacking = true;
		this.attackScheduler = this.attackTimer.scheduleAtFixedRate(this.attack, 0, this.attackSpeed, TimeUnit.MILLISECONDS);
	}
	
	public void stopAttackTimer()
	{
		this.attackScheduler.cancel(false);
		this.attacking = false;
	}
	
	@Override
	public abstract String getName();

	@Override
	public int getLife() {
		return this.life;
	}
	
	@Override
	public boolean removeLife(int life) {
		boolean rval = false;
		
		this.life -= life;

		if (this.life <= 0)
		{
			rval = true;
			//Transmission.sendToMapInstance(this.getMapInstance(), this.getDeadPacket());
			this.clearAll();
		}
			
		return rval;
	}
	
	public abstract void clearAll(); // Une fois mort, cette méthode sera appelé pour effacer toutes les occurences de l'objet. A redéfinir donc.
	
	//public abstract OutputBuffer getDeadPacket();

	@Override
	public abstract int getDamage();
	
	@Override
	public void attack(IKillable target) {
		if (this.target != null && this.getMapInstance() == this.target.getMapInstance() && this.getPosition().distance(target.getPosition()) <= 1)
		{	
			if (target instanceof Player)
			{
				((Player)target).partyLock.readLock().lock();
				
				if (((Player)target).party != null)
				{
					((Player)target).party.membersLock.lock();
				}
			}
			
			target.prepareToFight();
			if (!target.isDead())
			{
				int damage = this.getDamage();
				Transmission.sendDamageDisplay(this, target, damage);
				if (target.removeLife(damage))
				{
					this.stopAttackTimer();
				}
			}
			else
			{
				this.stopAttackTimer();
			}
			target.escapeFromFight();
			
			if (target instanceof Player)
			{
				if (((Player)target).party != null)
				{
					((Player)target).party.membersLock.unlock();
				}
				
				((Player)target).partyLock.readLock().unlock();
			}
		}
		else
		{
			this.stopAttackTimer();
		}
		
	}
	
	public Directions getTargetDirection(Position target)
	{
		Directions rval = null;
		ArrayList<Position> path = this.getPath(target);
		
		if (path.size() > 1 ) // Un chemin existe
		{
			// Le dernier élément est la position du MapFighter donc l'enlever pour obtenir la suivante
			path.remove(path.size()-1);
			if (path.get(path.size()-1).getX() < this.getX())
			{
				rval= Directions.DIR_LEFT;
			}
			
			if (path.get(path.size()-1).getX() > this.getX())
			{
				rval = Directions.DIR_RIGHT;
			}
			
			if (path.get(path.size()-1).getY() < this.getY())
			{
				rval = Directions.DIR_UP;
			}
			
			if (path.get(path.size()-1).getY() > this.getY())
			{
				rval = Directions.DIR_DOWN;
			}
		}
		
		return rval;
	}
	
	public void move() // Synchronisé avec Update de Npc
	{
		this.getMapInstance().lock.lock();
		this.prepareToFight(); // To not be killed during a move
		if (!this.isDead())
		{
			if (this.target != null)
			{
				int distance = this.getPosition().distance(this.target.getPosition());	

				if (distance == 1) // Fin de mouvement car se trouve à côté de l'ennemi
				{
					this.stopMovementTimer();
					this.sendStopMovePacketToMap();
					
					// To test
					Directions targetDirection = this.getTargetDirection(this.target.getPosition());
					if (targetDirection != this.getDir())
					{
						this.setDir(targetDirection);
						Transmission.sendToMapInstance(this.getMapInstance(), this.getDirPacket());
					}
					// End to test
					
					if (!this.attacking)
					{
						this.launchAttackTimer();
					}
				}
				else
				{
					Directions targetDirection = this.getTargetDirection(this.target.getPosition());
					
					if (targetDirection == null)
					{
						this.stopMovementTimer();
						this.sendStopMovePacketToMap();
					}
					else
					{
						if (targetDirection != this.getDir())
						{
							this.setDir(targetDirection);
							Transmission.sendToMapInstance(this.getMapInstance(), this.getDirMovePacket());
						}
						
						this.moveOneOnDir();
					}
				}
			}
			else if (this.destination != null)
			{
				int distance = this.getPosition().distance(this.destination);
				
				if (distance == 0) // Fin de mouvement car se trouve sur la case
				{			
					this.stopMovementTimer();
					this.sendStopMovePacketToMap();
				}
				else
				{
					Directions targetDirection = this.getTargetDirection(this.destination);
					if (targetDirection == null)
					{
						this.stopMovementTimer();
						this.sendStopMovePacketToMap();
					}
					else 
					{
						if (targetDirection != this.getDir())
						{
							this.setDir(targetDirection);
							Transmission.sendToMapInstance(this.getMapInstance(), this.getDirMovePacket());
						}
						
						this.moveOneOnDir();
					}
				}
			}
			else
			{
				this.stopMovementTimer();
				this.sendStopMovePacketToMap();
			}
		}
		this.escapeFromFight();
		this.getMapInstance().lock.unlock();
	}
	
	private class pathNode
	{
		public Position source;
		public int distanceParcourue;
		public int distanceRestante;
		
		public pathNode(Position source, int distanceParcourue, int distanceRestante)
		{
			this.source = source;
			this.distanceParcourue = distanceParcourue;
			this.distanceRestante = distanceRestante;
		}
		
		public int getPoid()
		{
			return this.distanceParcourue + this.distanceRestante;
		}
	}
	
	private void testPosition(Position positionToTest, Map.Entry<Position, pathNode> current, TreeMap<Position, pathNode> listeOuverte, TreeMap<Position, pathNode> listeFermee, Position destination)
	{
		if ((this.getMapInstance().getTileAllocation(positionToTest).isTraversableBy(this) || positionToTest.equals(destination)) && !listeFermee.containsKey(positionToTest) && (positionToTest.getX() >= 0 && positionToTest.getY() >= 0 && positionToTest.getX() < this.getMapInstance().getMap().mapAttributes.tiles.length && positionToTest.getY() < this.getMapInstance().getMap().mapAttributes.tiles[0].length))
		{
			if (!listeOuverte.containsKey(positionToTest))
			{
				listeOuverte.put(positionToTest, new pathNode(current.getKey(), current.getValue().distanceParcourue+1, positionToTest.distance(destination)));
			}
			else
			{
				if (current.getValue().distanceParcourue+1 < listeOuverte.get(positionToTest).distanceParcourue)
				{
					listeOuverte.put(positionToTest, new pathNode(current.getKey(), current.getValue().distanceParcourue+1, positionToTest.distance(destination)));
				}
			}
		}
	}
	
	private ArrayList<Position> getPath(Position destination)
	{
		ArrayList<Position> path = new ArrayList<Position>();
		
		TreeMap<Position, pathNode> listeOuverte = new TreeMap<Position, pathNode>();
		TreeMap<Position, pathNode> listeFermee = new TreeMap<Position, pathNode>();

		listeOuverte.put(this.getPosition(), new pathNode(null, 0, this.getPosition().distance(destination)));
		Map.Entry<Position, pathNode> current;
		do
		{
			current = listeOuverte.firstEntry();
			for (Map.Entry<Position, pathNode> entry : listeOuverte.entrySet())
			{
				if (entry.getValue().getPoid() < current.getValue().getPoid())
				{
					current = entry;
				}
			}
			
			listeFermee.put(current.getKey(), current.getValue());
			current = listeFermee.ceilingEntry(current.getKey()); // On récupère l'objet dans la liste fermée car en supprimant dans la liste ouverte on va le perdre (référence)
			listeOuverte.remove(current.getKey());

			this.testPosition(new Position((byte)(current.getKey().getX()+1), current.getKey().getY()), current, listeOuverte, listeFermee, destination);
			this.testPosition(new Position(current.getKey().getX(), (byte)(current.getKey().getY()+1)), current, listeOuverte, listeFermee, destination);
			this.testPosition(new Position((byte)(current.getKey().getX()-1), current.getKey().getY()), current, listeOuverte, listeFermee, destination);
			this.testPosition(new Position(current.getKey().getX(), (byte)(current.getKey().getY()-1)), current, listeOuverte, listeFermee, destination);

		} while (!current.getKey().equals(destination) && (listeOuverte.size() != 0));

		if (current.getKey().equals(destination))
		{
			path.add(destination);
			while (current.getValue().source != null)
			{
				path.add(current.getValue().source);
				current = listeFermee.ceilingEntry(current.getValue().source);
			}
		}
		
		return path;
	}
	
	public void attackTarget()
	{
		Runnable attack = new TalkingRunnable(new Runnable() {
			
			public void run() {
				if (MapFighter.this.tryToPrepareToFight())
				{
					if (!MapFighter.this.isDead())
					{
						if (MapFighter.this.target != null)
						{	
							MapFighter.this.destination = null;
							
							if (!MapFighter.this.moving)
							{
								int distance = MapFighter.this.getPosition().distance(MapFighter.this.target.getPosition());
								
								if (distance <= 1) // On es déjà a côté de l'ennemi, attaquer !
								{
									if (!MapFighter.this.attacking)
									{
										MapFighter.this.launchAttackTimer();
									}
								}
								else
								{	
									Directions direction = MapFighter.this.getTargetDirection(MapFighter.this.target.getPosition());
									if (direction != null)
									{
										MapFighter.this.setDir(direction);

										MapFighter.this.speed = 4;
										MapFighter.this.launchMovementTimer();
										
										Transmission.sendToMapInstance(MapFighter.this.getMapInstance(), MapFighter.this.getStartMovePacket());
									}
									
								}
							}
						}
					}
					MapFighter.this.escapeFromFight();
				}
			}
		});
		ServerConfiguration.getInstance().scheduledExecutor.submit(attack);
	}

	@Override
	public abstract OutputBuffer getStartMovePacket();

	@Override
	protected abstract OutputBuffer getStopMovePacket();

	@Override
	public abstract OutputBuffer getDirMovePacket();
	
	@Override
	public abstract OutputBuffer getDirPacket();
	
	@Override
	public boolean isDead()
	{
		return this.getLife() <= 0;
	}
	
	@Override
	public boolean tryToPrepareToFight()
	{
		return this.fightLock.tryLock();
	}
	
	@Override
	public void prepareToFight()
	{
		this.fightLock.lock();
	}
	
	@Override
	public void escapeFromFight()
	{
		this.fightLock.unlock();
	}
}
