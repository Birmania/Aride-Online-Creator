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

import java.io.DataInputStream;
import java.io.IOException;
import java.net.Socket;
import java.net.SocketException;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.concurrent.LinkedBlockingQueue;

import Communications.HandleData;
import Communications.OutputBuffer;
import Communications.Transmission;
import Exceptions.DoNotUseException;
import Miscs.MessageLogger;
import Miscs.TypeTools;


public class ClientThread implements Runnable {
	private short id;
	
	public Thread thread;
	private Socket socket;
	
	private LinkedBlockingQueue<OutputBuffer> packets;
	private Thread sender;
	
	public String login;
	
	public Player player;
	
	private boolean isInGame;
	
	
	
	public ClientThread(short id, Socket socket) {
		this.id = id;
		this.setLogin("");
		this.socket = socket;
		try {
			this.socket.setTcpNoDelay(true);
		} catch (SocketException e) {
			MessageLogger.getInstance().log(e);
		}
		
		this.packets = new LinkedBlockingQueue<OutputBuffer>();
		this.sender = new Thread()
        {
 
            public void run()
            {
                while (ClientThread.this.sender != null)
                {
                	try {
                		OutputBuffer packet = ClientThread.this.packets.take();

						try {
							// convert the buffer to a byte array
							byte packetArray[] = packet.toByteArray();
							
							// create a byte array containing the size of the main buffer
							byte nbByte[] = TypeTools.intToByteArray(Integer.reverseBytes(packetArray.length));
							
							// create the final packet
							byte finalPacket[] = new byte[packetArray.length+nbByte.length];
							System.arraycopy(nbByte, 0, finalPacket, 0, nbByte.length);
							System.arraycopy(packetArray, 0, finalPacket, nbByte.length, packetArray.length);
							
							// send the final packet
							ClientThread.this.socket.getOutputStream().write(finalPacket);
						} catch (IOException e) {

						}
					} catch (InterruptedException e) {

					}
                }
            }
            
            public void interrupt()
            {
            	ClientThread.this.sender = null;
            	super.interrupt();
            }
        };
		this.sender.start();
		
	    this.thread = new Thread(this); // instanciation du thread
	    this.isInGame = false;
	    
	    // Send the index to the player
	    OutputBuffer packet = new OutputBuffer("SYourIndex");
	    packet.writeShort(this.getId());
	    //this.sendPacket(packet);
	    Transmission.sendToClient(this, packet);
	}
	
	public boolean isInGame()
	{
		//return this.player != null;
		return this.isInGame;
	}
	
	public void setInGame(boolean value)
	{
		this.isInGame = value;
		
		/*if (value)
		{			

		}
		else
		{
			this.player.stopTiredness();
		}*/
	}
	
	public void setLogin(String login) {
		this.login = login;
	}
	
	public String getLogin() {
		return this.login;
	}
	
	public int receivePacket() {
		int packetId = 0;
		int nbBytes = 0;

		try
		{
			DataInputStream r = new DataInputStream(this.socket.getInputStream());
			// Read the number of bytes in the packet
			nbBytes = Integer.reverseBytes(r.readInt());
			if (nbBytes >= 4)
			{
				packetId = Integer.reverseBytes(r.readInt());
				byte buffer[] = new byte[nbBytes-4];
				r.read(buffer, 0, nbBytes-4);
				HandleData.getInstance().handle(this, packetId, buffer);
			}	
		}
		//Client s'est déconnecté, le thread doit s'arrêter et on sort le joueur de la population
		catch (IOException e)
		{
			Thread exit = new Exiter();
			exit.run();
			//this.exitClientSafe();
		}
		
		return packetId;
	}
	
	public void sendPacket(OutputBuffer packet) throws DoNotUseException
	{
		try {
			this.packets.put(packet);
		} catch (InterruptedException e) {
			MessageLogger.getInstance().log(e);
		}
	}
	
	class Exiter extends Thread {
		public void run() {			
			boolean mustSave = false;

			if (ClientThread.this.isInGame()) // Si un joueur est en train de jouer (et donc sur une map), on le retire de la map
			{
				mustSave = true;
				ClientThread.this.player.partyLock.writeLock().lock();
				ClientThread.this.player.prepareToFight(); // Pour ne pas faire de traitement en même temps que SGoBorderMap par exemple

				
				if (ClientThread.this.player.party != null)
				{
					ClientThread.this.player.party.removeMember(ClientThread.this.player);
				}
				
				ClientThread.this.player.quitMap();
					
				OutputBuffer packet = new OutputBuffer("SLeft");
				packet.writeShort(ClientThread.this.player.getId());
				Transmission.sendToMapInstanceBut(ClientThread.this.player.getMapInstance(), ClientThread.this.player, packet);
					
			
				ClientThread.this.setInGame(false);
				ClientThread.this.player.stopTiredness();
				
				for (Item currentItem : ClientThread.this.player.effectItems)
				{
					ClientThread.this.player.removeItemEffects(currentItem);
				}
				ClientThread.this.player.effectItems.clear();
				
				ClientThread.this.player.escapeFromFight();
				ClientThread.this.player.partyLock.writeLock().unlock();
			}
			
			ClientThread.this.thread = null; // On arrête le thread
			ClientThread.this.sender.interrupt();
			
			Population.getInstance().removePlayer(ClientThread.this, mustSave); // On retire aussi le joueur de la population totale	
		}
	}

	public void run()
	{
		while(this.thread != null) {
			this.receivePacket();
		}
	}
	
	public void LoadCharacter(ResultSet account) throws SQLException
	{
		this.player = new Player(this, World.getInstance().getMap(account.getShort("playerMap")).getOriginInstance(), account.getByte("PlayerX"), account.getByte("PlayerY"), account);
	}
	
	public short getId()
	{
		return this.id;
	}

	// TODO : devra être supprimé dans la version finale
	@Override
	protected void finalize() throws Throwable {
		super.finalize();
	}
}