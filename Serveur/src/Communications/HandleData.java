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

package Communications;


import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.concurrent.TimeUnit;

import Enumerations.Colors;
import Enumerations.Directions;
import Enumerations.ItemTypes;
import Enumerations.MapElementTypes;
import Exceptions.NoNpcException;
import Exceptions.NoPlayerException;
import Interfaces.IKillable;
import Main.ClientThread;
import Main.Craft;
import Main.Dream;
import Main.Faun;
import Main.Item;
import Main.ItemSlot;
import Main.Pet;
import Main.Player;
import Main.Population;
import Main.Position;
import Main.ServerConfiguration;
import Main.World;
import Miscs.MessageLogger;
import Miscs.TalkingRunnable;
import Npc.Npc;
import PMap.Map.MapInstance;
import PMap.MapBorder;
import PMap.MapElementList;
import PMap.MapItem;
import PMap.MapMissile;

public class HandleData {
	private static HandleData INSTANCE = null;
	
	private Method packetFunctions[]; // Functions launch by client packet
	public Map<String, Integer> serverPackets  = new TreeMap<String, Integer>();
	
	private HandleData ()
	{
		// Create server packet list
		String serverPacketList[] = {
				"SFindServer",
				"SAlertMsg",
				"SErrorLogin",
				"SEndWarp",
				"SCheckForMap",
				"SYourIndex",
				"SLife",
				"SStamina",
				"SSleep",
				"SExperience",
				"SNextLevel",
				"SPartyBars",
				"SPlayerSkills",
				"SPlayerCrafts",
				"SInventory",
				"SInventorySlot",
				"SWeaponSlot",
				"SArmorSlot",
			    "SHelmetSlot",
			    "SShieldSlot",
				"SMapData",
				//"SNpcData",
				"SPlayerStartInfos",
				"SPlayerPosition",
				"SRequestParty",
				"SJoinParty",
				"SLeaveParty",
				"SChatMsg",
				"SPlayerStartMove", // Begin a movement
				"SPlayerStopMove", // Stop a movement for player
				"SPlayerDirMove", // Change of direction during a movement
				"SPlayerDir", // Change just the direction
				"SNpcStartMove", // Start a movement for npc
				"SNpcStopMove", // Stop a movement for npc
				"SNpcDirMove", // Change of direction during a movement
				"SNpcDir", // Change just the direction
				"SPetStartMove", // Start a movement for pet
				"SPetStopMove",
				"SPetDirMove", // Change of direction during a movement
				"SPetDir", // Change just the direction
				"SLeft", // Leave the game
				"SQuitMap", // Leave the map
				"SPlayerDead",
				"SDamageDisplay",
				"SGetItemDisplay",
				"SMapNpcData",
				//"SMapPetData",
				"SNpcDead",
				"SMissileAppear",
				"SMissileDisappear",
				"SSpawnMapItem",
				"SDeleteMapItem",
				"SPlayerMsg",
				"SStatistics",
				"SPetDead",
				"SAreaWeather",
				"STime",
				"SConfirmUseItem",
				"SCancelUseItem"
				};
		
		int i = 0;
		for (String packetName : serverPacketList)
		{
			serverPackets.put(packetName, i);
			i++;
		}
		
		// Create client packet list and associate them functions
		String clientPacketList[] = {
				"CFindServer",
				"CLogin",
				"CNeedMap",
				"CRequestParty",
				"CJoinParty",
				"CLeaveParty",
				"CPlayerMove",
				"CPlayerStopMove",
				"CPlayerDirMove",
				"CPlayerDir",
				"CAttack",
				"CFire",
				"CMapGetItem",
				"CSayMsg",
				"CGoBorderMap",
				"CPlayerSleep",
				"CUseItem",
				"CDropItem",
				"CTakeOutWeapon",
				"CTakeOutArmor",
				"CTakeOutHelmet",
				"CTakeOutShield",
				"CMoveInventoryItem",
				"CAddStr",
				"CAddDef",
				"CAddDex",
				"CAddSci",
				"CAddLang",
				"CUseSkill",
				"CAdopt",
				"CAbandon",
				"CExecuteCraft"
				};
		
		packetFunctions = new Method[clientPacketList.length];
		
		i = 0;
		for (String packetName : clientPacketList)
		{
			try {
				packetFunctions[i] = this.getClass().getMethod(packetName, ClientThread.class, InputBuffer.class);
				i++;
			} catch (NoSuchMethodException | SecurityException e1) {
				MessageLogger.getInstance().log(e1);
				//e1.printStackTrace();
			}
		}
	}
	
	public synchronized final static HandleData getInstance()
	{
		if (INSTANCE == null)
		{
			INSTANCE = new HandleData();
		}
		return INSTANCE;
	}

	public void handle(ClientThread client, int packetId, byte buffer[]) throws IOException
	{	
		try {
			if (packetId >= 0 && packetId < packetFunctions.length)
			{
				packetFunctions[packetId].invoke(this, client, new InputBuffer(buffer));
			}
			else
			{
				MessageLogger.getInstance().log("Packet ID out of bound. Packet ID : "+packetId);
			}
		} catch (IllegalAccessException | IllegalArgumentException
				| InvocationTargetException e) {
			MessageLogger.getInstance().log(e);
		}
	}

	public void CFindServer(ClientThread client, InputBuffer buffer) {
		OutputBuffer packet = new OutputBuffer("SFindServer");
		//packet.writeShort(buffer.readShort());
		packet.writeShort(Population.getInstance().getNbPlayers());
		//packet.writeShort(ServerConfiguration.getInstance().getMaxPlayers());

		//client.packets.put(packet);
		//client.sendPacket(packet);
		Transmission.sendToClient(client, packet);
	}
	
	// This method must be synchronized because we are testing double launch of same account
	public synchronized void CLogin(ClientThread client, InputBuffer buffer) {
		String login = buffer.readString();
		String password = buffer.readString();
		String clientVersion = buffer.readInt()+"."+buffer.readInt()+"."+buffer.readInt();
		
		if (ServerConfiguration.secCode1.equals(buffer.readString()) && ServerConfiguration.secCode2.equals(buffer.readString()) && ServerConfiguration.secCode3.equals(buffer.readString()) && ServerConfiguration.secCode4.equals(buffer.readString()))
		{
			if (!ServerConfiguration.getInstance().acceptedClientVersion.equals(clientVersion))
			{
				Transmission.sendAlertMsg(client, "Vous n'avez pas la bonne version du client");
			}
			else
			{
				client.setLogin(login);
				try {
					// Verify that the player is not logged
					boolean find = false;
					// TODO : Maybe look for a better structure to do not lock the population "so much time"
					Population.getInstance().playersLock.readLock().lock();
						Iterator<ClientThread> ite = Population.getInstance().getPlayers().values().iterator();
						while (ite.hasNext() && !find)
						{
							ClientThread currentClient = ite.next();

							if (currentClient.getLogin().equals(login) && currentClient != client)
							{
								find = true;
							}
						}
						
					if (!find)
					{
						Player playerFromCache = Population.getInstance().retrievePlayer(login);
						
						if (playerFromCache != null)
						{
							playerFromCache.client = client;
							client.player = playerFromCache;
							client.player.enterGame();
						}
						else
						{
							Connection con = ServerConfiguration.getInstance().getConnection();
							ResultSet account = ServerConfiguration.getInstance().sendSelectQuery(con, "SELECT * FROM Account WHERE playerEmail='"+login+"' AND playerPassword = '"+password+"';");

							if (account.next())
							{	
								client.LoadCharacter(account);
						
								// Inventory must be the first send because it can have multiple impact (max life, etc..)
								client.player.enterGame();

							}
							else
							{
								Transmission.sendErrorLogin(client, "Login ou mot de passe incorrect");
							}
							ServerConfiguration.getInstance().releaseConnection(con);
						}
					}
					else
					{
						Transmission.sendErrorLogin(client, "Ce compte est déjà présent dans le jeu");
					}
					Population.getInstance().playersLock.readLock().unlock();
				} catch (SQLException e) {
					MessageLogger.getInstance().log(e);
					Transmission.sendErrorLogin(client, "Erreur côté serveur : Re-essayer de vous connecter ou reportez nous l'erreur si elle persiste.");
				}
			}
		}
	}
	
	public void CNeedMap(ClientThread client, InputBuffer buffer) {
		// Enter in the map
		if (buffer.readByte() == 1)
		{
			client.player.getMapInstance().getMap().sendToPlayer(client.player);
		}
		client.player.enterMapInstance();

		OutputBuffer packet = new OutputBuffer("SEndWarp");
		Transmission.sendToClient(client, packet);

		client.player.warpLock.release();
	}
	
	public void CRequestParty(ClientThread client, InputBuffer buffer) {
		try {
			Player targetPlayer = Population.getInstance().getPlayer(buffer.readShort()).player;

			if (client.player != targetPlayer) // Check hackers
			{
				// Assurer un lock ordonné
				if (targetPlayer.getId() < client.player.getId())
				{
					targetPlayer.partyLock.writeLock().lock();
					client.player.partyLock.writeLock().lock();
				}
				else
				{
					client.player.partyLock.writeLock().lock();
					targetPlayer.partyLock.writeLock().lock();
				}
			
				if (targetPlayer.party == null || !targetPlayer.party.haveMember(targetPlayer))
				{	
					if (client.player.party == null) // Le créateur du groupe n'a pas encore de groupe
					{
						client.player.party = Population.getInstance().createParty(buffer.readString());
						client.player.party.addMember(client.player);
					}
					targetPlayer.party = client.player.party; // On set le groupe pour le receveur mais on ne l'ajoutera dans le groupe que si il acceptera l'invitation
					
					OutputBuffer packet = new OutputBuffer("SRequestParty");
					
					packet.writeShort(client.player.getId());
					
					Transmission.sendToClient(targetPlayer.client, packet);
				}
				else
				{
					Transmission.sendChatMsgToPlayer(client.player, "Le joueur "+targetPlayer.getName()+" est déjà dans un groupe.", Colors.White);
				}

				client.player.partyLock.writeLock().unlock();
				targetPlayer.partyLock.writeLock().unlock();
			}
		} catch (NoPlayerException e) {
			MessageLogger.getInstance().log(e);
		}
	}
	
	public void CJoinParty(ClientThread client, InputBuffer buffer) {
		client.player.partyLock.readLock().lock();
		if (client.player.party != null)
		{
			client.player.party.addMember(client.player);
		}
		client.player.partyLock.readLock().unlock();
	}
	
	public void CLeaveParty(ClientThread client, InputBuffer buffer) {
		client.player.partyLock.writeLock().lock();
		if (client.player.party != null)
		{
			client.player.party.removeMember(client.player);
		}
		client.player.partyLock.writeLock().unlock();
	}
	
	public void CPlayerMove(ClientThread client, InputBuffer buffer) {
		if (System.currentTimeMillis() >= client.player.nextMovement)
		{
			client.player.prepareToFight();
			if (!client.player.isDead())
			{
				if (client.player.movementController == buffer.readByte())
				{
					if (client.player.movementController == 1)
					{
						client.player.movementController = 0; // go back to prediction
					}
					
					byte clientX = buffer.readByte();
					byte clientY = buffer.readByte();
					if (clientX == client.player.getX() && clientY == client.player.getY()) // Déplacement uniquement si le joueur est sur la bonne position de départ. Cas 1
					{	
						client.player.speed = buffer.readByte();
						client.player.launchMovementTimer();
						
						// Envoi du déplacement à tous les joueurs
						Transmission.sendToMapInstanceBut(client.player.getMapInstance(), client.player, client.player.getStartMovePacket());
					}
				}
			}
			client.player.escapeFromFight();
		}
		else
		{
			client.player.sendStopMovePacketToPlayer();
		}
	}
	
	public void CPlayerStopMove(ClientThread client, InputBuffer buffer) {

		client.player.getMapInstance().lock.lock();

		client.player.prepareToFight();
		if (!client.player.isDead())
		{
			if (client.player.moving)
			{
				client.player.stopMovementTimer();
				// TODO : Etre sur que le timer est bien stope car si on lock et que de son cote il demande un lock par la suite...

				byte clientX = buffer.readByte();
				byte clientY = buffer.readByte();
				

				int difference =  Math.abs(clientX - client.player.getX())+Math.abs(clientY - client.player.getY());
				if (difference == 0) // Perfect match
				{
					client.player.sendStopMovePacketToMapBut();
				}
				else if (difference == 1) // Maybe a lazy packet
				{
					// Look if the new position of the player is not allocated. If not, OK. Else, TP on the real player position
					if (client.player.getMapInstance().getTileAllocation(clientX, clientY).isTraversableBy(client.player))
					{
						// Re-placer le joueur en faisant confiance au client. (On pourra éventuellement le changer en faisant des burst refresh initié par le serveur selon son gameAI)
						if (client.player.moveOnPosition(clientX, clientY) <= 1)
						{
							client.player.sendStopMovePacketToMapBut();
						}
					}
					else
					{
						// TODO : C'est souvent dans le cas où deux personnes se suivent. Où veulent s'arrêter à la même case. On peut éventuellement regarder le offset et téléporter celui qui a le plus petit offset
						client.player.sendStopMovePacketToMap();
					}
				}
				else // Teleport the client
				{
					client.player.sendStopMovePacketToMap();
				}
			}
		}
		client.player.escapeFromFight();
		client.player.getMapInstance().lock.unlock();
	}
	
	public void CPlayerDirMove(ClientThread client, InputBuffer buffer) {
		client.player.getMapInstance().lock.lock();
		client.player.prepareToFight();
		if (!client.player.isDead())
		{
			if (client.player.moving)
			{
				// Pause the movement to pause it in GameAI
				client.player.stopMovementTimer();
				// TODO : Pareil que TODO de Stop movement
				byte x = buffer.readByte();
				byte y = buffer.readByte();
				Directions dir = Directions.values()[buffer.readByte()];
				
				client.player.setDir(dir); // Direction
				
				int difference =  Math.abs(x - client.player.getX())+Math.abs(y - client.player.getY());
				if (difference == 0 || difference == 1) // Perfect or quasi match
				{
					if (client.player.moveOnPosition(x, y) <= 1) // New X, Y
					{	
						// L'envoi doit se faire avant le déplacement dans la direction
						Transmission.sendToMapInstanceBut(client.player.getMapInstance(), client.player, client.player.getDirMovePacket());
	
						client.player.launchMovementTimer();
					}
					else
					{
						Transmission.sendToMapInstanceBut(client.player.getMapInstance(), client.player, client.player.getDirPacket());
					}
				}
				else
				{
					client.player.sendStopMovePacketToMap();
					// essayons cette ligne
					Transmission.sendToMapInstanceBut(client.player.getMapInstance(), client.player, client.player.getDirPacket());
				}
			}
			else // Peut arriver dans le cas où le start move a été refusé. On va tout de même accepter la direction mais pas le déplacement
			{
				buffer.skipBytes(2); // On ignore deux  bytes (qui sont les x et y)
				
				this.CPlayerDir(client, buffer);
			}
		}
		client.player.escapeFromFight();
		client.player.getMapInstance().lock.unlock();
	}
	
	public void CPlayerDir(ClientThread client, InputBuffer buffer) {
		client.player.prepareToFight();
		if (!client.player.isDead())
		{
			if (!client.player.moving)
			{
				client.player.setDir(Directions.values()[buffer.readByte()]);
			
				Transmission.sendToMapInstanceBut(client.player.getMapInstance(), client.player, client.player.getDirPacket());
			}
		}
		client.player.escapeFromFight();
	}
	
	public void CAttack(ClientThread client, InputBuffer buffer) {
		if (System.currentTimeMillis() >= client.player.nextAttack)
		{
			client.player.nextMovement = System.currentTimeMillis() + (client.player.attackSpeed / 2);
			client.player.nextAttack = System.currentTimeMillis() + client.player.attackSpeed;
			
			byte targetType = buffer.readByte();
			short targetIndex = buffer.readShort();
			
			client.player.inventoryLock.lock();
			try {
				client.player.tryToAttack(MapElementTypes.values()[targetType], targetIndex);
			} catch (NoNpcException | NoPlayerException e) {
				MessageLogger.getInstance().log(e);
			}
			client.player.inventoryLock.unlock();
		}
	}
	
	public void CFire(ClientThread client, InputBuffer buffer) {
		if (System.currentTimeMillis() >= client.player.nextAttack)
		{
			client.player.nextMovement = System.currentTimeMillis() + (client.player.attackSpeed / 2);
			client.player.nextAttack = System.currentTimeMillis() + client.player.attackSpeed;
			client.player.inventoryLock.lock();

			if (client.player.getWeaponSlot().getItemId() >= 0)
			{
				Item weapon = World.getInstance().items.get(client.player.getWeaponSlot().getItemId());
				if (weapon.type == ItemTypes.ItemTypeMissile.getCode()) // Arme à distance
				{
					int iSlot = client.player.getInventoryPosition(weapon.datas[3]);
					if (iSlot != -1)
					{
						new MapMissile(client.player.getMapInstance(), client.player.getX(), client.player.getY(), client.player.getDir(), (byte)World.getInstance().items.get(client.player.getWeaponSlot().getItemId()).datas[2], client.player, client.player.getDamage());
						client.player.deleteItem((byte)iSlot, (short)1);
					}
				}
				else if (weapon.type == ItemTypes.ItemTypeThrowable.getCode())
				{
					new MapMissile(client.player.getMapInstance(), client.player.getX(), client.player.getY(), client.player.getDir(), (byte)World.getInstance().items.get(client.player.getWeaponSlot().getItemId()).datas[2], client.player, client.player.getDamage());
					client.player.getWeaponSlot().reduceItemVal((short)1);
					Transmission.sendPlayerWeapon(client.player);
				}
			}
			client.player.inventoryLock.unlock();
		}
	}
	
	public void CMapGetItem(ClientThread client, InputBuffer buffer) {
		byte x = buffer.readByte();
		byte y = buffer.readByte();
		
		synchronized(client.player.getMapInstance().items)
		{
			MapItem takenItem = client.player.getMapInstance().getFirstItem(new Position(x, y));
			
			if (takenItem != null)
			{
				if (client.player.addItem(takenItem.itemSlot))
				{
					client.player.getMapInstance().deleteItem(takenItem);

					OutputBuffer packet = new OutputBuffer("SGetItemDisplay");
					
					packet.writeShort(takenItem.itemSlot.getItemId());
					packet.writeShort(takenItem.itemSlot.getItemVal());
					
					Transmission.sendToClient(client, packet);
				}
			}
		}
	}
	
	public void CSayMsg(ClientThread client, InputBuffer buffer) {
		OutputBuffer packet = new OutputBuffer("SPlayerMsg");
		packet.writeShort(client.getId());
		packet.writeString(buffer.readString());
		
		Transmission.sendToMapInstance(client.player.getMapInstance(), packet);
	}
	
	public void CGoBorderMap(ClientThread client, InputBuffer buffer) {
		MapInstance ancientMap = client.player.getMapInstance();
		ancientMap.lock.lock();
		client.player.prepareToFight();
		if (!client.player.isDead())
		{	
			byte XOrY = buffer.readByte();
			
			byte x = 0;
			byte y = 0;
			switch(client.player.getDir())
			{
			case DIR_UP:
				x = XOrY;
				break;
			case DIR_DOWN:
				x = XOrY;
				y = client.player.getMapInstance().getMaxY();
				break;
			case DIR_LEFT:
				y = XOrY;
				break;
			case DIR_RIGHT:
				x = client.player.getMapInstance().getMaxX();
				y = XOrY;
				break;
			}
			
			PMap.Map.MapInstance map = null;
			for (MapBorder currentBorder : client.player.getMapInstance().getMap().mapAttributes.borders)
			{
				if (currentBorder.XSource == x && currentBorder.YSource == y && currentBorder.DirectionSource == client.player.getDir().getCode())
				{
					if (client.player.dreamInstance != null)
					{
						map = client.player.dreamInstance.mapsInstances.get(currentBorder.mapDestination);
					}
					else
					{
						map = World.getInstance().getMap(currentBorder.mapDestination).getOriginInstance();
					}
					
					if (client.player.warpLock.tryAcquire())
					{
						client.player.warp(map, currentBorder.XDestination, currentBorder.YDestination);
					}
					
					break;
				}
			}
			
		}
		client.player.escapeFromFight();
		
		ancientMap.lock.unlock();
	}
	
	public void CPlayerSleep(ClientThread client, InputBuffer buffer) {
		MapInstance ancientMap = client.player.getMapInstance();
		ancientMap.lock.lock();
		client.player.prepareToFight();
		if (!client.player.isDead())
		{
			if (client.player.dreamInstance != null)
			{
				Transmission.sendChatMsgToPlayer(client.player, "Vous êtes déjà dans un rêve.", Colors.Red);
			}
			else
			{
				short dreamNum = (short)(Math.random()*(World.getInstance().dreams.size()));
				
				Dream selectedDream = World.getInstance().dreams.get(dreamNum);
				if (selectedDream == null)
				{
					MessageLogger.getInstance().log("Tentative de chargement du rêve : "+dreamNum+" mais inexistant.");
				}
				else
				{
					try
					{
						client.player.mapBeforeDream = client.player.getMapInstance();
						client.player.positionBeforeDream = new Position(client.player.getPosition());
						client.player.dreamInstance = selectedDream.newInstance();
						
						try {
							client.player.warpLock.acquire();
						} catch (InterruptedException e) {
							throw new Error();
						}
						client.player.warp(client.player.dreamInstance.mapsInstances.get(selectedDream.beginningMap), (byte)selectedDream.beginningX, (byte)selectedDream.beginningY);
					}
					catch (NullPointerException e)
					{
						MessageLogger.getInstance().log("Le rêve : "+dreamNum+" tente de charger une carte qui n'existe pas.");
					}
				}
			}
		}
		client.player.escapeFromFight();
		ancientMap.lock.unlock();
	}
	
	public void CUseItem(ClientThread client, InputBuffer buffer) {
		byte invSlotNum = buffer.readByte();
		
		// Avoid deadlock
		client.player.partyLock.readLock().lock();
		if (client.player.party != null)
		{
			client.player.party.membersLock.lock();
		}
		// End avoid
		
		// Prepare player to fight because equip an item could reduce player life
		client.player.prepareToFight();
		if (!client.player.isDead())
		{
			client.player.inventoryLock.lock();// On synchronize pour eviter les sauvegarde d'inventaire en plein milieu de la tâche

			ItemSlot usedSlot = new ItemSlot(client.player.inventory[invSlotNum]);
			if (usedSlot.getItemId() >= 0)
			{
				if (ItemTypes.isWeapon(usedSlot.getItemId()))
				{
					client.player.inventory[invSlotNum].setItemId(client.player.getWeaponSlot().getItemId());
					client.player.inventory[invSlotNum].addItemVal((short)1);
					client.player.inventory[invSlotNum].setItemDur(client.player.getWeaponSlot().getItemDur());
					client.player.setWeaponSlot(usedSlot);
					
					Transmission.sendPlayerWeapon(client.player);
					Transmission.sendPlayerInventorySlot(client.player, invSlotNum);
				}
				else if (ItemTypes.isArmor(usedSlot.getItemId()))
				{
					client.player.inventory[invSlotNum].setItemId(client.player.getArmorSlot().getItemId());
					client.player.inventory[invSlotNum].addItemVal((short)1);
					client.player.inventory[invSlotNum].setItemDur(client.player.getArmorSlot().getItemDur());
					client.player.setArmorSlot(usedSlot);
					
					Transmission.sendPlayerArmor(client.player);
					Transmission.sendPlayerInventorySlot(client.player, invSlotNum);
				}
				else if (ItemTypes.isHelmet(usedSlot.getItemId()))
				{
					client.player.inventory[invSlotNum].setItemId(client.player.getHelmetSlot().getItemId());
					client.player.inventory[invSlotNum].addItemVal((short)1);
					client.player.inventory[invSlotNum].setItemDur(client.player.getHelmetSlot().getItemDur());
					client.player.setHelmetSlot(usedSlot);
					
					Transmission.sendPlayerHelmet(client.player);
					Transmission.sendPlayerInventorySlot(client.player, invSlotNum);
				}
				else if (ItemTypes.isShield(usedSlot.getItemId()))
				{
					client.player.inventory[invSlotNum].setItemId(client.player.getShieldSlot().getItemId());
					client.player.inventory[invSlotNum].addItemVal((short)1);
					client.player.inventory[invSlotNum].setItemDur(client.player.getShieldSlot().getItemDur());
					client.player.setShieldSlot(usedSlot);
					
					Transmission.sendPlayerShield(client.player);
					Transmission.sendPlayerInventorySlot(client.player, invSlotNum);
				}
				else if (ItemTypes.isPotion(usedSlot.getItemId()))
				{
					Item potion = World.getInstance().items.get(usedSlot.getItemId());
					synchronized(client.player.effectItems)
					{
						if (!client.player.effectItems.contains(potion))
						{
							int duration = potion.datas[0];

							if (duration > 0)
							{
								client.player.effectItems.add(potion);
								
								final ClientThread theClient = client;
								final short itemId = usedSlot.getItemId();
								
								Runnable remover = new TalkingRunnable(new Runnable() {
									
									public void run() {
										theClient.player.partyLock.readLock().lock();
										if (theClient.player.party != null)
										{
											theClient.player.party.membersLock.lock();
										}
										
										theClient.player.prepareToFight(); // On ne teste pas si le joueur est mort car dans tous les cas il faut retirer l'effet

										if (theClient.isInGame())
										{
											Item potion = World.getInstance().items.get(itemId);
											
											theClient.player.removeItemEffects(potion);
											theClient.player.effectItems.remove(potion);
											
											OutputBuffer packet = new OutputBuffer("SCancelUseItem");
											packet.writeShort(theClient.player.getId());
											packet.writeShort(itemId);
											Transmission.sendToPartyElseClient(theClient, packet);
											
											
										}
										theClient.player.escapeFromFight();

										if (theClient.player.party != null)
										{
											theClient.player.party.membersLock.unlock();
										}
										theClient.player.partyLock.readLock().unlock();
									}
								});
								ServerConfiguration.getInstance().scheduledExecutor.schedule(remover, duration, TimeUnit.MILLISECONDS);
								
							}
							
							client.player.inventory[invSlotNum].reduceItemVal((short)1);
					
							client.player.applyItemEffects(potion);
							
							OutputBuffer packet = new OutputBuffer("SConfirmUseItem");
							packet.writeShort(client.player.getId());
							packet.writeShort(usedSlot.getItemId());
							Transmission.sendToPartyElseClient(client, packet);
							
							// Pourrait presque etre mis cote client avec le paquet SConfirmUseItem
							Transmission.sendPlayerInventorySlot(client.player, invSlotNum);
						}
					}
				}
			}
			//}
			client.player.inventoryLock.unlock();
		}
		client.player.escapeFromFight();
		
		// avoid deadlock
		if (client.player.party != null)
		{
			client.player.party.membersLock.unlock();
		}
		client.player.partyLock.readLock().unlock();
	}
	
	public void CDropItem(ClientThread client, InputBuffer buffer) {
		byte invSlotNum = buffer.readByte();
		short amount = buffer.readShort();
		
		client.player.inventoryLock.lock();

		ItemSlot usedSlot = client.player.inventory[invSlotNum];
		if (usedSlot.getItemId() >= 0)
		{
			if (usedSlot.getItemVal() >= amount)
			{
				client.player.getMapInstance().addItem(new MapItem(new Position(client.player.getPosition()), new ItemSlot(usedSlot.getItemId(), amount, usedSlot.getItemDur())));
				usedSlot.reduceItemVal(amount);
				Transmission.sendPlayerInventorySlot(client.player, invSlotNum);
			}
		}
		client.player.inventoryLock.unlock();
	}
	
	public void CTakeOutWeapon(ClientThread client, InputBuffer buffer) {
		// avoid deadlock
		client.player.partyLock.readLock().lock();
		if (client.player.party != null)
		{
			client.player.party.membersLock.lock();
		}
		
		client.player.prepareToFight();
		if (!client.player.isDead())
		{
			client.player.inventoryLock.lock();
			if (client.player.getWeaponSlot().getItemId() >= 0)
			{
				if (client.player.addItem(client.player.getWeaponSlot())) // Implique un sendplayerinventoryslot si success
				{
					client.player.setWeaponSlot(new ItemSlot((short)-1, (short)-1, (short)-1));
					Transmission.sendPlayerWeapon(client.player);
				}
			}
			client.player.inventoryLock.unlock();
		}
		client.player.escapeFromFight();
		
		// Avoid deadlock
		if (client.player.party != null)
		{
			client.player.party.membersLock.unlock();
		}
		client.player.partyLock.readLock().unlock();
	}
	
	public void CTakeOutArmor(ClientThread client, InputBuffer buffer) {
		// avoid deadlock
		client.player.partyLock.readLock().lock();
		if (client.player.party != null)
		{
			client.player.party.membersLock.lock();
		}
		
		client.player.prepareToFight();
		if (!client.player.isDead())
		{
			client.player.inventoryLock.lock();

			if (client.player.getArmorSlot().getItemId() >= 0)
			{
				if (client.player.addItem(client.player.getArmorSlot())) // Implique un sendplayerinventoryslot si success
				{
					client.player.setArmorSlot(new ItemSlot((short)-1, (short)-1, (short)-1));
					Transmission.sendPlayerArmor(client.player);
				}
			}
			
			client.player.inventoryLock.unlock();
		}
		client.player.escapeFromFight();
		
		// Avoid deadlock
		if (client.player.party != null)
		{
			client.player.party.membersLock.unlock();
		}
		client.player.partyLock.readLock().unlock();
	}
	
	public void CTakeOutHelmet(ClientThread client, InputBuffer buffer) {
		// avoid deadlock
		client.player.partyLock.readLock().lock();
		if (client.player.party != null)
		{
			client.player.party.membersLock.lock();
		}
		
		client.player.prepareToFight();
		if (!client.player.isDead())
		{
			client.player.inventoryLock.lock();

			if (client.player.getHelmetSlot().getItemId() >= 0)
			{
				if (client.player.addItem(client.player.getHelmetSlot())) // Implique un sendplayerinventoryslot si success
				{
					client.player.setHelmetSlot(new ItemSlot((short)-1, (short)-1, (short)-1));
					Transmission.sendPlayerHelmet(client.player);
				}
			}

			client.player.inventoryLock.unlock();
		}
		client.player.escapeFromFight();
		
		// Avoid deadlock
		if (client.player.party != null)
		{
			client.player.party.membersLock.unlock();
		}
		client.player.partyLock.readLock().unlock();
	}
	
	public void CTakeOutShield(ClientThread client, InputBuffer buffer) {
		// avoid deadlock
		client.player.partyLock.readLock().lock();
		if (client.player.party != null)
		{
			client.player.party.membersLock.lock();
		}
		
		// On bloque le joueur car désequipper un objet peut jouer sur la vie
		client.player.prepareToFight();
		if (!client.player.isDead())
		{
			client.player.inventoryLock.lock();

			if (client.player.getShieldSlot().getItemId() >= 0)
			{
				if (client.player.addItem(client.player.getShieldSlot())) // Implique un sendplayerinventoryslot si success
				{
					client.player.setShieldSlot(new ItemSlot((short)-1, (short)-1, (short)-1));
					Transmission.sendPlayerShield(client.player);
				}
			}

			client.player.inventoryLock.unlock();
		}
		client.player.escapeFromFight();
		
		// Avoid deadlock
		if (client.player.party != null)
		{
			client.player.party.membersLock.unlock();
		}
		client.player.partyLock.readLock().unlock();
	}
	
	public void CMoveInventoryItem(ClientThread client, InputBuffer buffer) {
		byte sourceInvSlot = buffer.readByte();
		byte targetInvSlot = buffer.readByte();
		short itemId = buffer.readShort();
		short itemVal = buffer.readShort();
		
		client.player.inventoryLock.lock();
	
		if (client.player.inventory[sourceInvSlot].getItemId() == itemId && client.player.inventory[sourceInvSlot].getItemVal() >= itemVal)
		{
			if (client.player.inventory[targetInvSlot].getItemId() == -1)
			{
				client.player.inventory[targetInvSlot].setItemId(itemId);
				client.player.inventory[targetInvSlot].addItemVal(itemVal);
				client.player.inventory[targetInvSlot].setItemDur(client.player.inventory[sourceInvSlot].getItemDur());
				client.player.inventory[sourceInvSlot].reduceItemVal(itemVal);
			}
			else if (client.player.inventory[targetInvSlot].getItemId() == itemId && World.getInstance().items.get(itemId).empilable == 1)
			{
				client.player.inventory[targetInvSlot].addItemVal(itemVal);
				client.player.inventory[sourceInvSlot].reduceItemVal(itemVal);
			}
			else
			{
				ItemSlot temp = client.player.inventory[targetInvSlot];
				client.player.inventory[targetInvSlot] = client.player.inventory[sourceInvSlot];
				client.player.inventory[sourceInvSlot] = temp;
			}
			Transmission.sendPlayerInventorySlot(client.player, targetInvSlot);
			Transmission.sendPlayerInventorySlot(client.player, sourceInvSlot);
		}

		client.player.inventoryLock.unlock();
	}
	
	public void CAddStr(ClientThread client, InputBuffer buffer) {
		
		client.player.statisticLock.lock();

		if (client.player.freePoints > 0)
		{
			client.player.strength++;
			client.player.freePoints--;
		}
		Transmission.sendPlayerStatistics(client.player);
		
		client.player.statisticLock.unlock();
	}
	
	public void CAddDef(ClientThread client, InputBuffer buffer) {
		client.player.statisticLock.lock();
		
		if (client.player.freePoints > 0)
		{
			client.player.defense++;
			client.player.freePoints--;
		}
		Transmission.sendPlayerStatistics(client.player);
		
		client.player.statisticLock.unlock();
	}
	
	public void CAddDex(ClientThread client, InputBuffer buffer) {
		client.player.statisticLock.lock();

		if (client.player.freePoints > 0)
		{
			client.player.dexterity++;
			client.player.freePoints--;
		}
		Transmission.sendPlayerStatistics(client.player);
		
		client.player.statisticLock.unlock();
	}
	
	public void CAddSci(ClientThread client, InputBuffer buffer) {
		client.player.statisticLock.lock();

		if (client.player.freePoints > 0)
		{
			client.player.science++;
			client.player.freePoints--;
		}
		Transmission.sendPlayerStatistics(client.player);

		client.player.statisticLock.unlock();
	}
	
	public void CAddLang(ClientThread client, InputBuffer buffer) {
		client.player.statisticLock.lock();

		if (client.player.freePoints > 0)
		{
			client.player.language++;
			client.player.freePoints--;
		}
		Transmission.sendPlayerStatistics(client.player);
		
		client.player.statisticLock.unlock();
	}
	
	public void CUseSkill(ClientThread client, InputBuffer buffer) {
		if (System.currentTimeMillis() >= client.player.nextAttack)
		{
			client.player.skillsLock.lock();
			
			client.player.nextMovement = System.currentTimeMillis() + (client.player.attackSpeed / 2);
			client.player.nextAttack = System.currentTimeMillis() + client.player.attackSpeed;
		
			short skillNum = buffer.readShort();
			
			if (World.getInstance().skills.get(skillNum).target == 1)
			{
				byte x = buffer.readByte();
				byte y = buffer.readByte();
				MapElementList elements = client.player.getMapInstance().getAreaAllocation(x, y, World.getInstance().skills.get(skillNum).range);
				

				for (IKillable currentElement : elements.getIKillable())
				{
					
					client.player.attack(currentElement, 20);
					
				}
			}
			else if (World.getInstance().skills.get(skillNum).target == 0)
			{
				byte targetType = buffer.readByte();
				short targetId = buffer.readShort();
				
				client.player.petLock.lock();
				if (client.player.pet != null)
				{
					if (targetType==MapElementTypes.Npc.getCode())
					{
						client.player.pet.prepareToFight();
						if (!client.player.pet.isDead())
						{
							try {
								client.player.pet.isFollowing = false;
								client.player.pet.target = client.player.getMapInstance().getNpc(targetId);
								
								client.player.pet.attackTarget();
								
							} catch (NoNpcException e) {
								MessageLogger.getInstance().log(e);
							}
								
						}
						client.player.pet.escapeFromFight();
					}
				}
				client.player.petLock.unlock();
			}
			else if (World.getInstance().skills.get(skillNum).target == 2)
			{
				client.player.petLock.lock();
				if (client.player.pet != null)
				{
					client.player.pet.followMaster();
				}
				client.player.petLock.unlock();
			}
			
			client.player.skillsLock.unlock();
		}
	}
	
	public void CAdopt(ClientThread client, InputBuffer buffer) {
		try {
			Population.getInstance().playersLock.readLock().lock();
			synchronized(client.player.getMapInstance().getPlayers())
			{
				client.player.petLock.lock();
				if (client.player.pet == null)
				{
					Npc adoptedNpc = client.player.getMapInstance().getNpc(buffer.readShort());
					adoptedNpc.prepareToFight();
					if (!adoptedNpc.isDead())
					{
						
						adoptedNpc.destroy();
						client.player.pet = new Pet(client.getId(), client.player.getMapInstance(), adoptedNpc.getX(), adoptedNpc.getY(), Faun.getInstance().faun[adoptedNpc.type.id]);
						
						client.player.pet.setXY(adoptedNpc.getPosition());
						client.player.getMapInstance().addObserver(client.player.pet);
						
						OutputBuffer packet = new OutputBuffer("SPlayerStartInfos");
						packet.writeByte((byte)1);
						packet.writeShort(client.player.getId());
						client.player.writeStartInfosInPacket(packet);
						
						Transmission.sendToAllPlayer(packet);
						
						packet = new OutputBuffer("SPlayerPosition");
						packet.writeShort((short)1);
						client.player.writePositionInPacket(packet);
						Transmission.sendToMapInstance(client.player.getMapInstance(), packet);
						
						
					}
					adoptedNpc.escapeFromFight();
				}
				client.player.petLock.unlock();
			}
			Population.getInstance().playersLock.readLock().unlock();
		} catch (NoNpcException e) {
			MessageLogger.getInstance().log(e);
		}
	}
	
	public void CAbandon(ClientThread client, InputBuffer buffer) {
		client.player.petLock.lock();
		if (client.player.pet != null)
		{
			Pet playerPet = client.player.pet;
			playerPet.prepareToFight();
			
			if (!playerPet.isDead())
			{
				synchronized(playerPet.getMapInstance().getNpcs())
				{
					Set<Short> npcIds = playerPet.getMapInstance().getNpcs().keySet();
					short newId = (short)playerPet.getMapInstance().getMap().mapAttributes.npcs.length;
					while (npcIds.contains(newId))
					{
						newId++;
					}
					playerPet.getMapInstance().addNpc(new Npc(newId, playerPet.getMapInstance(), playerPet.getX(), playerPet.getY(), playerPet.type));
					
					playerPet.clearAll();
					client.player.pet = null;
				}
			}
			
			playerPet.escapeFromFight();
		}
		client.player.petLock.unlock();
	}
	
	public void CExecuteCraft(ClientThread client, InputBuffer buffer) {
		boolean allMaterials = true;
		Craft craftToExecute = World.getInstance().crafts.get(buffer.readShort());
		
		client.player.inventoryLock.lock();
		int iSlot;
		// Vérifier la présence des matériaux
		for (Craft.Material currentMaterial : craftToExecute.materials)
		{
			iSlot = client.player.getInventoryPosition(currentMaterial.itemId);
			if (iSlot > -1)
			{
				if (client.player.inventory[iSlot].getItemVal() < currentMaterial.itemCount)
				{
					allMaterials = false;
					break;
				}
			}
			else
			{
				allMaterials = false;
				break;
			}
		}
		
		if (allMaterials)
		{
			for (Craft.Material currentMaterial : craftToExecute.materials)
			{
				iSlot = client.player.getInventoryPosition(currentMaterial.itemId);
				client.player.inventory[iSlot].reduceItemVal(currentMaterial.itemCount);
				Transmission.sendPlayerInventorySlot(client.player, (byte)iSlot);
			}
			
			for (Craft.Material currentProduct : craftToExecute.products)
			{
				client.player.addItem(new ItemSlot(currentProduct.itemId, currentProduct.itemCount, (short)10));
			}
		}
		client.player.inventoryLock.unlock();
	}
}
