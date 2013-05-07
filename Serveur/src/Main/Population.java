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

import java.net.Socket;
import java.sql.Connection;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.TreeMap;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.locks.ReadWriteLock;
import java.util.concurrent.locks.ReentrantReadWriteLock;

import Exceptions.NoPlayerException;
import Interfaces.IRecognizable;
import Miscs.MessageLogger;
import Miscs.TalkingRunnable;


public class Population {
	private static Population INSTANCE = null;
	
	private TreeMap<Short, ClientThread> players;
	private List<ClientThread> playersToSave;
	public ReadWriteLock playersLock;
	public TreeMap<Short, Party> partys;
	
	private Population() {
		this.players = new TreeMap<Short, ClientThread>();
		this.playersToSave = new ArrayList<ClientThread>();
		this.playersLock = new ReentrantReadWriteLock(true);
		
		// Creating partys
		this.partys = new TreeMap<Short, Party>();
		
		ScheduledExecutorService saver = Executors.newSingleThreadScheduledExecutor();
		saver.scheduleAtFixedRate(new TalkingRunnable(new SaveAllPlayers()), 900, 900, TimeUnit.SECONDS);
	}
	
	public synchronized final static Population getInstance()
	{
		if (INSTANCE == null)
		{
			INSTANCE = new Population();
		}
		return INSTANCE;
	}
	
	public void addPlayer(Socket socket)
	{
		//synchronized(this.players) // Permet à ce que la population de joueurs ne soit pas itérée pendant qu'on la modifie (et inversement)
		//{
		Population.getInstance().playersLock.writeLock().lock();
		short i = 1;
		while (this.players.containsKey(i))
		{
			i++;
		}
		
		if (i < ServerConfiguration.getInstance().maxPlayers)
		{
			ClientThread newPlayer = new ClientThread(i, socket);

			this.players.put(i, newPlayer);
			newPlayer.thread.start(); 	//TODO : Découvrir le bug de nullpointer exception sur cette ligne
		}
		//}
		Population.getInstance().playersLock.writeLock().unlock();
	}
	
	public void removePlayer(ClientThread client, boolean mustSave)
	{
		Population.getInstance().playersLock.writeLock().lock();
		//synchronized(this.players) // Permet à ce que la population de joueurs ne soit pas itérée pendant qu'on la modifie (et inversement)
		//{
		if (mustSave)
		{
			//client.player.stopTiredness();
			//client.setInGame(false);
			this.addToSave(client);
		}
		
		this.players.remove(client.getId());
		
		//}
		Population.getInstance().playersLock.writeLock().unlock();
	}
	
	public ClientThread getPlayer(int index) throws NoPlayerException
	{
		ClientThread client = this.players.get((short)index);
		if (client == null)
		{
			throw new NoPlayerException();
		}
		return client;
	}
	
	public TreeMap<Short, ClientThread> getPlayers()
	{
		return this.players;
	}
	
	public short getNbPlayers()
	{
		return (short)this.players.size();
	}
	
	public Party createParty(String partyName)
	{
		Party p = new Party(partyName);
		
		p.setId((short)-1);
		short i = 0;
		while ((p.getId() == -1) && (i < this.partys.size()+1)) // deuxième condition non obligatoire
		{
			if (!this.partys.containsKey(i))
			{
				p.setId(i);
				this.partys.put(i, p);
			}
			i++;
		}
		
		return p;
	}
	
	private void addToSave(ClientThread client)
	{
		synchronized(this.playersToSave)
		{
			this.playersToSave.add(client);
		}
	}
	
	public Player retrievePlayer(String login)
	{
		Player rval = null;
		
		synchronized(this.playersToSave)
		{
			for (ClientThread currentClient : this.playersToSave)
			{
				if (currentClient.login.equals(login))
				{
					rval = currentClient.player;
					this.playersToSave.remove(currentClient);
					break;
				}
			}
		}
		
		return rval;
	}
	
	class SaveAllPlayers implements Runnable {

		@Override
		public void run() {
			String saveQuery = "";
			/*synchronized(Population.this.getPlayers())
			{*/
			Population.getInstance().playersLock.readLock().lock();
			synchronized(Population.getInstance().playersToSave)
			{
				List<ClientThread> toSave = Population.getInstance().playersToSave;
				for (ClientThread currentClient : Population.this.getPlayers().values())
				{
					if (currentClient.isInGame())
					{
						toSave.add(currentClient);
					}
				}
				
				Collections.sort(toSave, new IRecognizable.IRecognizableComparator());
				
				for (ClientThread currentClient : toSave)
				{
					currentClient.player.prepareToSave();
				}
				
				for (ClientThread currentClient : toSave)
				{
					saveQuery += currentClient.player.savePlayer();
				}
				
				for (ClientThread currentClient : toSave)
				{
					currentClient.player.releaseFromSave();
				}
				//}
				
				Population.getInstance().playersToSave.clear();
			}
			Population.getInstance().playersLock.readLock().unlock();
			if (saveQuery != "")
			{
				try {
					Connection con = ServerConfiguration.getInstance().getConnection();
					ServerConfiguration.getInstance().sendUpdateQuery(con, saveQuery);
					ServerConfiguration.getInstance().releaseConnection(con);
				} catch (SQLException e) {
					MessageLogger.getInstance().log(e);
				}
			}
		}
		
	}
}
