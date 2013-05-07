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
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Observable;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.locks.Lock;
import java.util.concurrent.locks.ReentrantLock;

import Communications.InputBuffer;
import Communications.OutputBuffer;
import Communications.Transmission;
import Enumerations.TileTypes;
import Exceptions.NoNpcException;
import Exceptions.NoValidSpawnPositionException;
import Interfaces.IKillable;
import Interfaces.IRecognizable;
import Main.Faun;
import Main.Pet;
import Main.Player;
import Main.Position;
import Main.ServerConfiguration;
import Miscs.TalkingRunnable;
import Npc.Npc;
import Npc.NpcAttributes;
import Tile.TileBlock;


public class Map{
	public short id;
	
	public String md5;
	
	// Map attributes
	public MapAttributes mapAttributes;
	
	private ArrayList<MapInstance> instances;
	
	private static void processConflict(ArrayList<IKillable> allKillable, ArrayList<MapMissile> allMapMissile)
	{
		Collections.sort(allMapMissile, new IRecognizable.IRecognizableComparator());
		Collections.sort(allKillable, new IRecognizable.IRecognizableComparator());
		
		for (IKillable currentKillable : allKillable)
		{
			if (currentKillable instanceof Player)
			{
				((Player)currentKillable).partyLock.readLock().lock();
				
				if (((Player)currentKillable).party != null)
				{
					((Player)currentKillable).party.membersLock.lock();
				}
			}
		}
		
		for (MapMissile currentMapMissile : allMapMissile)
		{
			currentMapMissile.lockMovement();
		}
		
		// Done by attack
		/*for (IKillable currentKillable : allKillable)
		{
			currentKillable.prepareToFight();
		}*/
		
		for (IKillable currentKillable : allKillable)
		{
			for (MapMissile currentMapMissile : allMapMissile)
			{
				if (currentKillable.isDead()) break;
				
				if (currentMapMissile.moving)
				{
					if (currentKillable != currentMapMissile.getLauncher())
					{
						if (currentKillable instanceof Player)
						{
							if ((((Player)currentKillable).party == null) || !((Player)currentKillable).party.haveMember((Player)currentMapMissile.getLauncher()))
							{
								currentMapMissile.attack(currentKillable);
							}
						}
						else
						{
							currentMapMissile.attack(currentKillable);
						}
					}
				}
			}
		}
		
		/*for (IKillable currentKillable : allKillable)
		{
			currentKillable.escapeFromFight();
		}*/
		
		for (MapMissile currentMapMissile : allMapMissile)
		{
			currentMapMissile.unlockMovement();
		}
		
		for (IKillable currentKillable : allKillable)
		{
			if (currentKillable instanceof Player)
			{
				if (((Player)currentKillable).party != null)
				{
					((Player)currentKillable).party.membersLock.unlock();
				}
				
				((Player)currentKillable).partyLock.readLock().unlock();
			}
		}
	}
	
	public Map(short numMap, String md5, InputBuffer mapBuffer)
	{	
		this.id = numMap;
		this.md5 = md5;
		
		this.instances = new ArrayList<MapInstance>();
		
		this.mapAttributes = new MapAttributes();
		this.mapAttributes.deserialize(mapBuffer);
	}
	
	public void sendToPlayer(Player player)
	{
		OutputBuffer mapPacket = new OutputBuffer("SMapData");
		mapPacket.writeShort(this.id);
		this.mapAttributes.writeInBuffer(mapPacket);

		//player.sendPacket(mapPacket);
		Transmission.sendToClient(player.client, mapPacket);
	}
	
	public MapInstance newInstance()
	{
		MapInstance newMI = new MapInstance();
		
		this.instances.add(newMI);
		
		return newMI;
	}
	
	// TODO : A supprimer pour faire du traitement plus intelligent
	public MapInstance getOriginInstance()
	{
		return this.instances.get(0);
	}
	
	public byte getMaxX()
	{
		return (byte)(this.mapAttributes.tiles.length - 1);
	}
	
	public byte getMaxY()
	{
		return (byte)(this.mapAttributes.tiles[0].length - 1);
	}
	
	public class MapInstance extends Observable{
		// Map variables
		private ArrayList<Player> players; // Joueurs sur la carte
		
		private HashMap<Short, Npc> npcs;
		
		private MapElementList tilesAllocation[][]; // Cases de la carte pour savoir si elle est libre ou pas
		
		public HashMap<Byte, MapMissile> missiles;
		
		public ArrayList<MapItem> items;
		
		public Lock lock;
		
		public MapInstance()
		{
			this.players = new ArrayList<Player>();
			this.npcs = new HashMap<Short, Npc>();
			this.missiles = new HashMap<Byte, MapMissile>();
			this.items = new ArrayList<MapItem>();
			
			// Init the tile allocation
			this.tilesAllocation = new MapElementList[Map.this.mapAttributes.tiles.length][Map.this.mapAttributes.tiles[0].length];
			
			for (int i = 0 ; i < this.tilesAllocation.length ; i ++)
			{
				for (int j = 0 ; j < this.tilesAllocation[i].length ; j++)
				{
					this.tilesAllocation[i][j] = new MapElementList();
				}
			}
			
			// Allocate specific tiles
			for (byte i = 0 ; i < Map.this.mapAttributes.tiles.length ; i++)
			{
				for (byte j = 0 ; j < Map.this.mapAttributes.tiles[i].length ; j++)
				{
					if (Map.this.mapAttributes.tiles[i][j].type == TileTypes.Blocked.getCode())
					{
						new TileBlock(this, i, j);
						//this.setTileAllocation(i, j, new TileBlock(this, i, j));
					}
				}
			}
			
			this.lock = new ReentrantLock();
			
			this.spawnNpcs();
		}
		
		public Map getMap()
		{
			return Map.this;
		}
		
		public byte getMaxX()
		{
			return Map.this.getMaxX();
		}
		
		public byte getMaxY()
		{
			return Map.this.getMaxY();
		}
		
		public Position getClosestSpawnPosition(MapMovable element) throws NoValidSpawnPositionException
		{
			Position position = element.getPosition();
			
			int maxTry = (this.getMaxX()+1)*(this.getMaxY()+1);
			
			int nbTry = 0;
			int generalOffset = 1;
			int XOffset = 0;
			int YOffset = 0;
			while (!isValidSpawnPosition(element, position.getX()+XOffset, position.getY()+YOffset))
			{
				if (mapContainPosition(position.getX()+XOffset, position.getY()+YOffset))
				{
					nbTry++;
					if (nbTry == maxTry)
					{
						throw new NoValidSpawnPositionException();
					}
				}
				
				if (XOffset == -1 && YOffset == -generalOffset)
				{
					generalOffset++;
					YOffset = -generalOffset;
					XOffset = 0;
				}
				else
				{
					if (YOffset == -generalOffset)
					{
						if (XOffset == generalOffset)
						{
							YOffset++;
						}
						else
						{
							XOffset++;
						}
					}
					else if (YOffset == generalOffset)
					{
						if (XOffset == -generalOffset)
						{
							YOffset--;
						}
						else
						{
							XOffset--;
						}
					}
					else
					{
						if (XOffset > 0)
						{
							YOffset++;
						}
						else
						{
							YOffset--;
						}
					}
				}
			}
			
			return new Position((byte)(position.getX()+XOffset), (byte)(position.getY()+YOffset));
		}
		
		private Boolean isValidSpawnPosition(MapMovable element, int X, int Y)
		{
			return this.mapContainPosition(X, Y) && this.getTileAllocation((byte)X, (byte)Y).isTraversableBy(element);
		}
		
		private Boolean mapContainPosition(int X, int Y)
		{
			return X >= 0 && X <= this.getMaxX()
					&& Y >= 0 && Y <= this.getMaxY();
		}
		
		public void checkConflict(byte X, byte Y)
		{
			final ArrayList<IKillable> allKillable = this.getTileAllocation(X, Y).getIKillable();
			final ArrayList<MapMissile> allMapMissile = this.getTileAllocation(X, Y).getMapMissile();
			
			Runnable task = new TalkingRunnable(new Runnable() {
				@Override
				public void run() {
		            Map.processConflict(allKillable, allMapMissile);
				}
			});
			ServerConfiguration.getInstance().scheduledExecutor.submit(task);
		}
		
		public void addPlayer(Player player, byte x, byte y)
		{
			synchronized(this.players)
			{
				if (!this.players.contains(player))
				{
					this.players.add(player);
					
					// Set the tiles to occupy
					player.setXY(x, y);
					
					player.petLock.lock();
					if (player.pet != null)
					{
						//player.pet.setXY(player.pet.getPosition());
						// Set X et Y avant le setXY car le remove pourrait se faire sur une position hors map sinon
						player.pet.getPosition().setX(player.getX());
						player.pet.getPosition().setY(player.getY());
						player.pet.setXY(player.getPosition());
						player.getMapInstance().addObserver(player.pet);
					}
					player.petLock.unlock();
				}
			}
		}
		
		public void removePlayer(Player player)
		{
			synchronized(this.players)
			{
				this.players.remove(player);
				
				this.removeTileAllocation(player.getPosition(), player);
				player.petLock.lock();
				if (player.pet != null)
				{
					this.removePet(player.pet);
				}
				
				player.petLock.unlock();
			}
		}
		
		public void removePet(Pet pet)
		{
			this.removeTileAllocation(pet.getPosition(), pet);
		}
		
		public ArrayList<Player> getPlayers()
		{
			return this.players;
		}
		
		public MapElementList getTileAllocation(byte x, byte y)
		{
			MapElementList rval = new MapElementList();
			
			if (x >= 0 && x < Map.this.mapAttributes.tiles.length && y >= 0 && y < Map.this.mapAttributes.tiles[0].length) // x et y = case valide
			{
				rval = tilesAllocation[x][y];
			}
			return rval;
		}
		
		public MapElementList getTileAllocation(Position position)
		{
			return this.getTileAllocation(position.getX(), position.getY());
		}
		
		public MapElementList getAreaAllocation(byte x, byte y, short range)
		{
			MapElementList rval = new MapElementList();
			int offsetX = -range;
			int offsetY = -range;
			
			this.lock.lock();
			
			while (x+offsetX <= x+range)
			{
				while(y+offsetY <= y+range)
				{
					rval.add(this.getTileAllocation((byte)(x+offsetX), (byte)(y+offsetY)));
					offsetY++;
				}
				offsetY = -range;
				offsetX++;
			}
			
			this.lock.unlock();
			return rval;
		}
		
		public void addTileAllocation(byte x, byte y, MapElement value)
		{
			tilesAllocation[x][y].add(value);
	
			this.setChanged();
			this.notifyObservers(value);
		}
		
		public void addTileAllocation(Position position, MapElement value)
		{
			this.addTileAllocation(position.getX(), position.getY(), value);
		}
		
		public void removeTileAllocation(byte x, byte y, MapElement value)
		{
			this.tilesAllocation[x][y].remove(value);
		}
		
		public void removeTileAllocation(Position position, MapElement value)
		{
			this.removeTileAllocation(position.getX(), position.getY(), value);
		}
		
		public MapElement getElement(byte x, byte y)
		{
			MapElement rval = null;
			
			return rval;
		}
		
		public Npc getNpc(int index) throws NoNpcException
		{
			Npc npc = this.npcs.get((short)index);
			if (npc == null)
			{
				throw new NoNpcException();
			}
			return npc;
		}
		
		public HashMap<Short, Npc> getNpcs()
		{
			return this.npcs;
		}
		
		public void spawnNpcs()
		{
			if (Map.this.mapAttributes.npcs != null)
			{
				synchronized(this.npcs)
				{
					
					int nbNpcs = Map.this.mapAttributes.npcs.length;
					
					//Npc newNpc;
					for(int i = 0 ; i < nbNpcs ; i++)
					{
						this.spawnNpc(i);
						/*NpcAttributes attrib = this.mapAttributes.npcs[i];
						if (attrib.hasardSpawn == true)
						{
							byte x;
							byte y;
							do
							{
								x = (byte)(Math.random()*(this.mapAttributes.tiles.length-1));
								y = (byte)(Math.random()*(this.mapAttributes.tiles[0].length-1));
							} while (this.getTileAllocation(x, y).size() > 0);
			
							newNpc = new Npc((short)i, this, x, y, Faun.getInstance().faun[this.mapAttributes.npcs[i].id]);
						}
						else
						{
							newNpc = new Npc((short)i, this, attrib.x[0], attrib.y[0], Faun.getInstance().faun[this.mapAttributes.npcs[i].id]);
						}
						this.addTileAllocation(newNpc.getPosition(), newNpc);
						this.npcs.put((short)i, newNpc);
						this.addObserver(newNpc);*/
					}
				}
			}
		}
		
		public void spawnNpc(int numNpc)
		{
			Npc newNpc;
			
			NpcAttributes attrib = Map.this.mapAttributes.npcs[numNpc];
			if (attrib.hasardSpawn == true)
			{
				byte x;
				byte y;
				do
				{
					x = (byte)(Math.random()*(Map.this.mapAttributes.tiles.length-1));
					y = (byte)(Math.random()*(Map.this.mapAttributes.tiles[0].length-1));
				} while (this.getTileAllocation(x, y).size() > 0);
	
				newNpc = new Npc((short)numNpc, this, x, y, Faun.getInstance().faun[Map.this.mapAttributes.npcs[numNpc].id]);
			}
			else
			{
				newNpc = new Npc((short)numNpc, this, attrib.x[0], attrib.y[0], Faun.getInstance().faun[Map.this.mapAttributes.npcs[numNpc].id]);
			}
			
			this.addNpc(newNpc);
			/*this.addTileAllocation(newNpc.getPosition(), newNpc);
			this.npcs.put((short)numNpc, newNpc);
			this.addObserver(newNpc);
			
			OutputBuffer packet = new OutputBuffer("SMapNpcData");
			packet.writeByte((byte)1);
			newNpc.writeInPacket(packet);
			
			synchronized(this.players)
			{
				Iterator<Player> ite = this.players.iterator();
				
				while (ite.hasNext())
				{
					ite.next().sendPacket(packet);
				}
			}*/
		}
		
		public void addNpc(Npc newNpc)
		{
			this.addTileAllocation(newNpc.getPosition(), newNpc);
			this.npcs.put(newNpc.getId(), newNpc);
			this.addObserver(newNpc);
			
			OutputBuffer packet = new OutputBuffer("SMapNpcData");
			packet.writeByte((byte)1);
			newNpc.writeInPacket(packet);
			
			synchronized(this.players)
			{
				Iterator<Player> ite = this.players.iterator();
				
				while (ite.hasNext())
				{
					//ite.next().sendPacket(packet);
					Transmission.sendToClient(ite.next().client, packet);
				}
			}
		}
		
		public void removeNpc(Npc npc)
		{
			synchronized(this.npcs)
			{
				this.removeTileAllocation(npc.getPosition(), npc);
				//setTileAllocation(npc.getPosition(), null);
	
				this.npcs.remove(npc.getId());
			}
		}
		
		public void scheduleRespawn(int numNpc, long delay)
		{
			ScheduledExecutorService respawnTimer;
			respawnTimer = Executors.newSingleThreadScheduledExecutor();
			respawnTimer.schedule(new TalkingRunnable(new Respawner(numNpc)), delay, TimeUnit.SECONDS);
		}
		
		class Respawner implements Runnable {
			private int numNpc;
			
			public Respawner(int numNpc)
			{
				this.numNpc = numNpc;
			}
			
			@Override
			public void run() {
				MapInstance.this.spawnNpc(this.numNpc);
			}
		}
		
		public void sendLivingNpcsTo(Player player)
		{
			synchronized(this.getNpcs())
			{
				OutputBuffer packet = new OutputBuffer("SMapNpcData");
				
				packet.writeByte((byte)this.getNpcs().size());
				for (Npc currentNpc : this.getNpcs().values())
				{
					currentNpc.writeInPacket(packet);
				}
				
				//player.sendPacket(packet);
				Transmission.sendToClient(player.client, packet);
			}
		}
		
		/*public void sendLivingPetsTo(Player player)
		{
			short nbPet = 0;
			synchronized(this.players)
			{	
				OutputBuffer tempPacket = new OutputBuffer("");
				for (Player currentPlayer : this.players)
				{
					if (currentPlayer.pet != null)
					{
						nbPet++;
						currentPlayer.pet.writeInPacket(tempPacket);
					}
				}
				
				if (nbPet > 0)
				{
					OutputBuffer packet = new OutputBuffer("SMapPetData");

					packet.writeShort(nbPet);
					packet.writePacket(tempPacket);
					
					player.sendPacket(packet);
				}
			}
		}*/
		
		public void sendLivingPlayersTo(Player player)
		{
			synchronized(this.players)
			{
				ArrayList<Lock> petLocks = new ArrayList<Lock>();
				
				OutputBuffer packet = new OutputBuffer("SPlayerPosition");

				packet.writeShort((short)this.players.size());
				
				Iterator<Player> ite = this.players.iterator();
				Player currentClient;
				while (ite.hasNext())
				{
					currentClient = ite.next();
					petLocks.add(currentClient.petLock);
					currentClient.petLock.lock();
					currentClient.writePositionInPacket(packet);
				}
				
				//player.sendPacket(packet);
				Transmission.sendToClient(player.client, packet);
				
				for (Lock currentLock : petLocks)
				{
					currentLock.unlock();
				}
			}
		}
		
		public void sendItemsTo(Player player)
		{
			synchronized(this.items)
			{
				OutputBuffer packet = new OutputBuffer("SSpawnMapItem");
				packet.writeShort((short)this.items.size());
	
				for (MapItem mapItem : this.items)
				{
					mapItem.writeInPacket(packet);
				}
				
				//player.sendPacket(packet);
				Transmission.sendToClient(player.client, packet);
			}
		}
		
		public void addItem(MapItem mapItem)
		{
			synchronized(this.items)
			{
				this.items.add(mapItem);
				
				OutputBuffer packet = new OutputBuffer("SSpawnMapItem");
				packet.writeShort((short)1);
				mapItem.writeInPacket(packet);
				
				Transmission.sendToMapInstance(this, packet);
			}
		}
		
		public MapItem getFirstItem(Position position)
		{
			MapItem currentItem = null;
			synchronized(this.items)
			{
				Iterator<MapItem> ite = this.items.iterator();
				
				boolean find = false;
				while (ite.hasNext() && !find)
				{
					currentItem = ite.next();
					if (currentItem.position.equals(position))
					{
						find = true;
					}
				}
			}
			return currentItem;
		}
		
		public void deleteItem(MapItem mapItem)
		{
			synchronized(this.items)
			{
				this.items.remove(mapItem);
				
				OutputBuffer packet = new OutputBuffer("SDeleteMapItem");
				//currentItem.writeInPacket(packet);
				packet.writeByte(mapItem.position.getX());
				packet.writeByte(mapItem.position.getY());
				packet.writeShort(mapItem.itemSlot.getItemVal());
				
				Transmission.sendToMapInstance(this, packet);
			}
		}
	}
}
