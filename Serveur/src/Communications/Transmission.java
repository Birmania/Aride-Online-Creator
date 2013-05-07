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

import java.util.ArrayList;
import java.util.Iterator;
import java.util.concurrent.locks.Lock;

import Enumerations.Colors;
import Enumerations.MapElementTypes;
import Exceptions.DoNotUseException;
import Exceptions.NoPlayerException;
import Interfaces.IFighter;
import Interfaces.IKillable;
import Main.ClientThread;
import Main.ItemSlot;
import Main.Party;
import Main.Pet;
import Main.Player;
import Main.Population;
import Main.World;
import Miscs.MessageLogger;
import PMap.Map;
import PMap.MapElement;

public class Transmission {
	private static Transmission INSTANCE = null;


	//private ExecutorService router;
	
	private Transmission()
	{

		//this.router = Executors.newSingleThreadExecutor();
	
	}
	
	public final static Transmission getInstance() // would be synchronize but more efficient not to
	{
		if (INSTANCE == null)
		{
			INSTANCE = new Transmission();
		}
		return INSTANCE;
	}
	

	
	public static void sendToAllPlayer(OutputBuffer packet)
	{
		ClientThread currentClient;
		Population.getInstance().playersLock.readLock().lock();
		Iterator<ClientThread> ite = Population.getInstance().getPlayers().values().iterator();
		
		while (ite.hasNext())
		{
			try {
				currentClient = ite.next();
				if (currentClient.isInGame())
				{
					currentClient.sendPacket(packet);
				}
			} catch (DoNotUseException e) {
				// Must never happen
				MessageLogger.getInstance().log(e);
			}
		}
		Population.getInstance().playersLock.readLock().unlock();
	}
	
	public static void sendToMapInstance(Map.MapInstance map, OutputBuffer packet)
	{	
		synchronized(map.getPlayers())
		{
			Iterator<Player> ite = map.getPlayers().iterator();
			
			while (ite.hasNext())
			{
				try {
					ite.next().client.sendPacket(packet);
				} catch (DoNotUseException e) {
					// Must never happen
					MessageLogger.getInstance().log(e);
				}
			}
		}
	}
	
	public static void sendToMapInstanceBut(Map.MapInstance map, Player player, OutputBuffer packet)
	{
		synchronized(map.getPlayers())
		{
			Iterator<Player> ite = map.getPlayers().iterator();
			
			Player currentPlayer;
			while (ite.hasNext())
			{
				currentPlayer = ite.next();
				if (currentPlayer.getId() != player.getId())
				{
					//currentPlayer.sendPacket(packet);
					//Transmission.sendToPlayer(currentPlayer, packet);
					try {
						currentPlayer.client.sendPacket(packet);
					} catch (DoNotUseException e) {
						// Must never happen
						MessageLogger.getInstance().log(e);
					}
				}
			}
		}
	}
	
	public static void sendToParty(Party party, OutputBuffer packet)
	{
		party.membersLock.lock();
		Iterator<Player> ite = party.members.values().iterator();
		
		Player current;
		while (ite.hasNext())
		{
			current = ite.next();
			//current.sendPacket(packet);
			//Transmission.sendToPlayer(current, packet);
			try {
				current.client.sendPacket(packet);
			} catch (DoNotUseException e) {
				// Must never happen
				MessageLogger.getInstance().log(e);
			}
		}
		//}
		party.membersLock.unlock();
	}
	
	public static void sendToPartyElseClient(ClientThread client, OutputBuffer packet)
	{
		client.player.partyLock.readLock().lock();
		if (client.player.party != null)
		{
			Transmission.sendToParty(client.player.party, packet);
		}
		else
		{
			//player.sendPacket(packet);

			//Transmission.sendToPlayer(player, packet);
			try {
				client.sendPacket(packet);
			} catch (DoNotUseException e) {
				// Must never happen
				MessageLogger.getInstance().log(e);
			}
		}
		client.player.partyLock.readLock().unlock();
	}
	
	public static void sendToClient(ClientThread client, OutputBuffer packet)
	{
		try {
			client.sendPacket(packet);
		} catch (DoNotUseException e) {
			// Must never happen
			MessageLogger.getInstance().log(e);
		}
	}
	
	public static void sendAlertMsg(ClientThread client, String msg)
	{
		OutputBuffer packet = new OutputBuffer("SAlertMsg");
		packet.writeString(msg);
		//player.sendPacket(packet);
		Transmission.sendToClient(client, packet);
	}
	
	public static void sendErrorLogin(ClientThread client, String msg)
	{
		OutputBuffer packet = new OutputBuffer("SErrorLogin");
		packet.writeString(msg);
		//player.sendPacket(packet);
		Transmission.sendToClient(client, packet);
	}

	public static void sendChatMsgToPlayer(Player player, String msg, Colors color)
	{
		OutputBuffer packet = new  OutputBuffer("SChatMsg");
		
		packet.writeString(msg);
		
		packet.writeShort(color.getCode());
	
		//player.sendPacket(packet);
		Transmission.sendToClient(player.client, packet);
	}
	
	public static void sendPlayerLife(Player player)
	{	
		OutputBuffer packet = new OutputBuffer("SLife");
		packet.writeShort(player.getId());
		player.writeLifeInPacket(packet);
		
		// TODO : Envoyer à la partie
		Transmission.sendToPartyElseClient(player.client, packet);
		/*if (player.party != null)
		{
			synchronized (player.party.members)
			{
				Iterator<Player> ite = player.party.members.values().iterator();
				
				Player current;
				while (ite.hasNext())
				{
					current = ite.next();
					current.sendPacket(packet);
				}
			}
		}
		else
		{
			player.sendPacket(packet);
		}*/
	}
	
	public static void sendPlayerStamina(Player player)
	{	
		// Send bars
		OutputBuffer packet = new OutputBuffer("SStamina");
		packet.writeShort(player.getId());
		player.writeStaminaInPacket(packet);
		
		// TODO : Envoyer à la partie
		Transmission.sendToPartyElseClient(player.client, packet);
		//Transmission.sendToPlayer(player, packet);
		/*if (player.party != null)
		{
			synchronized (player.party.members)
			{
				Iterator<Player> ite = player.party.members.values().iterator();
				
				Player current;
				while (ite.hasNext())
				{
					current = ite.next();
					current.sendPacket(packet);
				}
			}
		}
		else
		{
			player.sendPacket(packet);
		}*/
	}
	
	public static void sendPlayerSleep(Player player)
	{	
		// Send bars
		OutputBuffer packet = new OutputBuffer("SSleep");
		player.writeSleepInPacket(packet);
		
		//player.sendPacket(packet);
		Transmission.sendToClient(player.client, packet);
	}
	
	public static void sendPlayerNextLevel(Player player)
	{
		// Send bars
		OutputBuffer packet = new OutputBuffer("SNextLevel");
		packet.writeInt(3000);
		
		//player.sendPacket(packet);
		Transmission.sendToClient(player.client, packet);
	}
	
	public static void sendPlayerExperience(Player player)
	{
		// Send bars
		OutputBuffer packet = new OutputBuffer("SExperience");
		player.writeExperienceInPacket(packet);
		
		//player.sendPacket(packet);
		Transmission.sendToClient(player.client, packet);
	}
	
	public static void sendPlayerInventory(Player player)
	{
		OutputBuffer packet = new OutputBuffer("SInventory");
		
		// Write
		// TODO : N'envoyer que la durabilité ou la valeur et côté client en fonction de l'item comprendre ce que c'est
		
		
		player.writeEquipmentInPacket(packet);
		/*packet.writeShort(player.getArmorSlot().getItemId());
		packet.writeShort(player.getArmorSlot().getItemVal());
		packet.writeShort(player.getArmorSlot().getItemDur());
		packet.writeShort(player.getWeaponSlot().getItemId());
		packet.writeShort(player.getWeaponSlot().getItemVal());
		packet.writeShort(player.getWeaponSlot().getItemDur());
		packet.writeShort(player.getHelmetSlot().getItemId());
		packet.writeShort(player.getHelmetSlot().getItemVal());
		packet.writeShort(player.getHelmetSlot().getItemDur());
		packet.writeShort(player.getShieldSlot().getItemId());
		packet.writeShort(player.getShieldSlot().getItemVal());
		packet.writeShort(player.getShieldSlot().getItemDur());*/
		
		
		/*packet.writeShort(player.petSlot.getItemId());
		packet.writeShort(player.petSlot.getItemVal());
		packet.writeShort(player.petSlot.getItemDur());*/
		
		for(ItemSlot iSlot : player.inventory)
		{
			packet.writeShort(iSlot.getItemId());
			packet.writeShort(iSlot.getItemVal());
			packet.writeShort(iSlot.getItemDur());
		}
		
		// TODO : Envoyer à la partie
		//player.sendPacket(packet);
		Transmission.sendToClient(player.client, packet);
	}
	
	public static void sendPlayerSkills(Player player)
	{
		player.skillsLock.lock();
		OutputBuffer packet = new OutputBuffer("SPlayerSkills");
		
		byte nbSkill = 0;
		
		OutputBuffer tempPacket = new OutputBuffer("");
		
		for(Short skillNum : player.skills)
		{
			if (skillNum > -1)
			{
				tempPacket.writeShort(skillNum);
				nbSkill++;
			}
		}
		
		if (nbSkill > 0)
		{
			packet.writeByte(nbSkill);
			packet.writePacket(tempPacket);
			//player.sendPacket(packet);
			Transmission.sendToClient(player.client, packet);
		}
		player.skillsLock.unlock();
	}
	
	public static void sendPlayerSkill(Player player, short slotNum)
	{
		OutputBuffer packet = new OutputBuffer("SPlayerSkills");
		
		packet.writeByte((byte)1);
		packet.writeShort(slotNum);
		packet.writeShort(player.skills.get(slotNum));
		
		//player.sendPacket(packet);
		Transmission.sendToClient(player.client, packet);
	}
	
	public static void sendPlayerCrafts(Player player)
	{
		OutputBuffer packet = new OutputBuffer("SPlayerCrafts");
		
		byte nbCraft = 0;
		
		OutputBuffer tempPacket = new OutputBuffer("");
		
		for(Short craftNum : player.crafts)
		{
			if (craftNum > -1)
			{
				tempPacket.writeShort(craftNum);
				nbCraft++;
			}
		}
		
		if (nbCraft > 0)
		{
			packet.writeByte(nbCraft);
			packet.writePacket(tempPacket);
			//player.sendPacket(packet);
			Transmission.sendToClient(player.client, packet);
		}
	}
	
	public static void sendPlayerArmor(Player player)
	{
		OutputBuffer packet = new OutputBuffer("SArmorSlot");
	
		packet.writeShort(player.getId());
		packet.writeShort(player.getArmorSlot().getItemId());
		packet.writeShort(player.getArmorSlot().getItemVal());
		packet.writeShort(player.getArmorSlot().getItemDur());
		
		//player.sendPacket(packet);
		Transmission.sendToPartyElseClient(player.client, packet);
	}
	
	public static void sendPlayerWeapon(Player player)
	{
		OutputBuffer packet = new OutputBuffer("SWeaponSlot");
		
		packet.writeShort(player.getId());
		packet.writeShort(player.getWeaponSlot().getItemId());
		packet.writeShort(player.getWeaponSlot().getItemVal());
		packet.writeShort(player.getWeaponSlot().getItemDur());
		
		//player.sendPacket(packet);
		Transmission.sendToPartyElseClient(player.client, packet);
	}
	
	public static void sendPlayerHelmet(Player player)
	{
		OutputBuffer packet = new OutputBuffer("SHelmetSlot");
		
		packet.writeShort(player.getId());
		packet.writeShort(player.getHelmetSlot().getItemId());
		packet.writeShort(player.getHelmetSlot().getItemVal());
		packet.writeShort(player.getHelmetSlot().getItemDur());
		
		//player.sendPacket(packet);
		Transmission.sendToPartyElseClient(player.client, packet);
	}
	
	public static void sendPlayerShield(Player player)
	{
		OutputBuffer packet = new OutputBuffer("SShieldSlot");
		
		packet.writeShort(player.getId());
		packet.writeShort(player.getShieldSlot().getItemId());
		packet.writeShort(player.getShieldSlot().getItemVal());
		packet.writeShort(player.getShieldSlot().getItemDur());
		
		//player.sendPacket(packet);
		Transmission.sendToPartyElseClient(player.client, packet);
	}
	
	public static void sendPlayerInventorySlot(Player player, byte numSlot)
	{
		OutputBuffer packet = new OutputBuffer("SInventorySlot");
		
		packet.writeByte(numSlot);
		packet.writeShort(player.inventory[numSlot].getItemId());
		packet.writeShort(player.inventory[numSlot].getItemVal());
		packet.writeShort(player.inventory[numSlot].getItemDur());
		
		//player.sendPacket(packet);
		Transmission.sendToClient(player.client, packet);
	}
	
	public static void sendPlayerStatistics(Player player)
	{
		OutputBuffer packet = new OutputBuffer("SStatistics");
		
		packet.writeShort(player.strength);
		packet.writeShort(player.defense);
		packet.writeShort(player.dexterity);
		packet.writeShort(player.science);
		packet.writeShort(player.language);
		
		packet.writeShort(player.freePoints);
		
		//player.sendPacket(packet);
		Transmission.sendToClient(player.client, packet);
	}
	
	public static void sendDamageDisplay(IFighter source, IKillable target, int damage)
	{
		OutputBuffer packet = new OutputBuffer("SDamageDisplay");
		packet.writeByte(MapElementTypes.get((MapElement)source).getCode());
		packet.writeShort(source.getId());
		packet.writeByte(MapElementTypes.get((MapElement)target).getCode());
		packet.writeShort(target.getId());
		packet.writeInt(damage);
		
		/*if (source instanceof Player)
		{
			((Player)source).sendPacket(packet);
		}
		else if (source instanceof Pet)
		{
			try {
				Population.getInstance().getPlayer(source.getId()).sendPacket(packet);
			} catch (NoPlayerException e) {
				MessageLogger.getInstance().log(e);
			}
		}*/
		if (source instanceof Pet)
		{
			try {
				source = (IFighter)Population.getInstance().getPlayer(source.getId());
			} catch (NoPlayerException e) {
				MessageLogger.getInstance().log(e);
			}
		}
		
		if (target instanceof Pet)
		{
			try {
				target = (IKillable)Population.getInstance().getPlayer(target.getId());
			} catch (NoPlayerException e) {
				MessageLogger.getInstance().log(e);
			}
		}
		
		if (target instanceof Player)
		{
			((Player)target).partyLock.readLock().lock();
			if (((Player)target).party != null)
			{
				((Player)target).party.membersLock.lock();
			}
			Transmission.sendToPartyElseClient((((Player)target).client), packet);
		
			if (source != target)
			{
				if (source instanceof Player)
				{
					if (((Player)target).party != null)
					{
						((Player)target).party.membersLock.lock();
						if (!((Player)target).party.haveMember((Player)source))
						{
							//((Player)source).sendPacket(packet);
							Transmission.sendToClient(((Player)source).client, packet);
						}
					}
					else
					{
						Transmission.sendToClient(((Player)source).client, packet);
					}
				}
			}
		}
		else
		{
			Transmission.sendToClient(((Player)source).client, packet);
		}
		
		/*if (source != target)
		{
			if (source instanceof Player)
			{
				//((Player)target).sendPacket(packet);
				Transmission.sendToParty((Player)target, packet);
			}
			else
			{
				Transmission.sendToPlayer(((Player)source), packet);
			}
		}*/
		
		/*if (source instanceof Player)
		{
			((Player)source).partyLock.readLock().unlock();
		}*/
		if (target instanceof Player)
		{
			if (((Player)target).party != null)
			{
				((Player)target).party.membersLock.unlock();
			}
			((Player)target).partyLock.readLock().unlock();
		}
	}
	
	/*public static void sendWeatherToAll()
	{
		synchronized(World.getInstance().cartography)
		{
			Transmission.sendToAllPlayer(World.getInstance().getAreaWeatherPacket());
		}
	}*/
	
	public static void sendWeatherToPlayer(Player player)
	{
		synchronized(World.getInstance().cartography)
		{
			Transmission.sendToClient(player.client, World.getInstance().getAreaWeatherPacket());
		}
	}
	
	/*public static void sendTimeToAll()
	{
		World.getInstance().timeLock.lock();
		Transmission.sendToAllPlayer(World.getInstance().getTimePacket());
		World.getInstance().timeLock.unlock();
	}*/
	
	public static void sendTimeToPlayer(Player player)
	{
		World.getInstance().timeLock.lock();
		Transmission.sendToClient(player.client, World.getInstance().getTimePacket());
		World.getInstance().timeLock.unlock();
	}
	
	public static void communicateNewPlayer(Player player)
	{
		player.petLock.lock();
		OutputBuffer packetNewPlayer = new OutputBuffer("SPlayerStartInfos");
		packetNewPlayer.writeByte((byte)1);
		packetNewPlayer.writeShort(player.getId());
		player.writeStartInfosInPacket(packetNewPlayer);
		
		byte nbPlayer = 0;
		OutputBuffer tempPacket = new OutputBuffer("");
		Iterator<ClientThread> ite = Population.getInstance().getPlayers().values().iterator();
		ArrayList<Lock> petLocks = new ArrayList<Lock>();
		while (ite.hasNext())
		{
			ClientThread currentClient = ite.next();
			
			if (currentClient.isInGame())
			{
				petLocks.add(currentClient.player.petLock);
				currentClient.player.petLock.lock();
				
				nbPlayer++;
				tempPacket.writeShort(currentClient.player.getId());
				currentClient.player.writeStartInfosInPacket(tempPacket);
				if (currentClient.player != player)
				{
					Transmission.sendToClient(currentClient, packetNewPlayer);
				}
			}
		}
		

		OutputBuffer packet = new OutputBuffer("SPlayerStartInfos");
		packet.writeByte(nbPlayer);
		packet.writePacket(tempPacket);
		
		Transmission.sendToClient(player.client, packet);
		
		// free all lock on pet
		for (Lock currentLock : petLocks)
		{
			currentLock.unlock();
		}
		player.petLock.unlock();
	}
}