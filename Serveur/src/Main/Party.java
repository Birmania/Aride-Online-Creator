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

import java.util.Map;
import java.util.TreeMap;
import java.util.concurrent.locks.Lock;
import java.util.concurrent.locks.ReentrantLock;

import Communications.OutputBuffer;
import Communications.Transmission;

public class Party {
	private short id;
	private String name;
	public TreeMap<Short, Player> members;
	public Lock membersLock;
	
	public Party(String partyName){
		this.members = new TreeMap<Short, Player>();
		this.membersLock = new ReentrantLock();
		this.name = partyName;
	}
	
	public void addMember(Player player)
	{
		player.partyLock.readLock().lock();

		this.membersLock.lock();
		short i = (short)this.members.size();
		
		if (i < ServerConfiguration.getInstance().maxPartyMembers)
		{	
			OutputBuffer packet = new  OutputBuffer("SJoinParty");

			packet.writeShort(this.getId());
			
			Transmission.sendToClient(player.client, packet);
			
			// do not get life if in modification
			player.prepareToFight();
			player.inventoryLock.lock();
			OutputBuffer packetNewMember = new OutputBuffer("SPartyBars");
			packetNewMember.writeByte((byte)1);
			packetNewMember.writeShort(player.getId());
			player.writeEquipmentInPacket(packetNewMember);
			player.writeLifeInPacket(packetNewMember);
			player.writeStaminaInPacket(packetNewMember);
			
			for (Player currentPlayer : this.members.values())
			{
				// Envoi au membre courant de la vie du nouveau membre
				Transmission.sendToClient(currentPlayer.client, packetNewMember);
			}
			player.inventoryLock.unlock();
			player.escapeFromFight();
			
			for (Player currentPlayer : this.members.values())
			{
				currentPlayer.prepareToFight();
				currentPlayer.inventoryLock.lock();
			}
			
			packet = new OutputBuffer("SPartyBars");
			packet.writeByte((byte)this.members.size());
			for (Player currentPlayer : this.members.values())
			{
				packet.writeShort(currentPlayer.getId());
				currentPlayer.writeEquipmentInPacket(packet);
				currentPlayer.writeLifeInPacket(packet);
				currentPlayer.writeStaminaInPacket(packet);
			}
			
			// Envoie de la vie de l'équipe au nouveau membre
			Transmission.sendToClient(player.client, packet);
			
			for (Player currentPlayer : this.members.values())
			{
				currentPlayer.inventoryLock.unlock();
				currentPlayer.escapeFromFight();
			}
			
			this.members.put(i, player);

		}

		this.membersLock.unlock();
		player.partyLock.readLock().unlock();
	}
	
	public void removeMember(Player player)
	{
		player.partyLock.writeLock().lock();
		this.membersLock.lock();

		for (Map.Entry<Short, Player> entry : this.members.entrySet())
		{
			if (entry.getValue() == player)
			{	
				this.members.remove(entry.getKey());
				
				player.party = null;
				
				OutputBuffer packet = new OutputBuffer("SLeaveParty");
				
				packet.writeShort(player.getId());
				
				Transmission.sendToParty(this, packet);

				break;
			}
		}

		this.membersLock.unlock();
		player.partyLock.writeLock().unlock();
	}
	
	public boolean haveMember(Player player)
	{
		return this.members.containsValue(player);
	}
	
	public void setId(short id)
	{
		this.id = id;
	}
	
	public short getId()
	{
		return this.id;
	}
}
