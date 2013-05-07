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

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Set;
import java.util.TreeMap;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.locks.Lock;
import java.util.concurrent.locks.ReentrantLock;

import Communications.InputBuffer;
import Communications.OutputBuffer;
import Communications.Transmission;
import Miscs.MessageLogger;
import Miscs.TalkingRunnable;
import Miscs.TypeTools;
import PMap.Area;
import PMap.Map;


public class World {
	private static World INSTANCE = null;
	
	public boolean time;	// Vrai : Jour, Faux : Nuit
	public Lock timeLock;
	private TreeMap<Short, Map> environment;
	public ArrayList<Area> cartography;
	public ArrayList<Item> items;
	public ArrayList<Dream> dreams;
	public ArrayList<Skill> skills;
	public ArrayList<Craft> crafts;
	
	private World() {
		// Init variables
		this.time = true;
		this.timeLock = new  ReentrantLock();
		this.environment = new TreeMap<Short, Map>();
		this.cartography = new ArrayList<Area>();
		this.items = new ArrayList<Item>();
		this.dreams = new ArrayList<Dream>();
		this.skills = new ArrayList<Skill>();
		this.crafts = new  ArrayList<Craft>();
		
		short i;
		
		// Loading dreams
		File dreamFolder = new File(ServerConfiguration.serverPath, "dreams");
		for (i = 0 ; i <= ServerConfiguration.getInstance().maxDreams ; i++)
		{
			File dreamFile = new File(dreamFolder.getPath(), "Dream"+i+".aod");

			if (dreamFile.length() > Integer.MAX_VALUE)
			{
				MessageLogger.getInstance().log("Erreur : Fichier "+dreamFile.getAbsolutePath()+" trop grand.");
			}
			else
			{
				// Create a byte array
				byte fileBytes[] = new byte[(int)dreamFile.length()];
				
				try {
					if (dreamFile.exists())
					{
						FileInputStream fileStream = (new FileInputStream(dreamFile));

						// Sock the file content in a buffer
						fileStream.read(fileBytes, 0, fileBytes.length);
						
						// Read the file buffer
						InputBuffer fileBuffer = new InputBuffer(fileBytes);

						// Load the map
						this.dreams.add(i, new Dream(fileBuffer));
					}
				} catch (IOException e) {
					MessageLogger.getInstance().log(e);
				}
			}
		}
		
		// Loading spells
		File skillFolder = new File(ServerConfiguration.serverPath, "skills");
		for (i = 0 ; i <= ServerConfiguration.getInstance().maxSkills ; i++)
		{
			File skillFile = new File(skillFolder.getPath(), "Skill"+i+".aos");

			if (skillFile.length() > Integer.MAX_VALUE)
			{
				MessageLogger.getInstance().log("Erreur : Fichier "+skillFile.getAbsolutePath()+" trop grand.");
			}
			else
			{
				// Create a byte array
				byte fileBytes[] = new byte[(int)skillFile.length()];
				
				try {
					if (skillFile.exists())
					{
						FileInputStream fileStream = (new FileInputStream(skillFile));

						// Sock the file content in a buffer
						fileStream.read(fileBytes, 0, fileBytes.length);
						
						// Read the file buffer
						InputBuffer fileBuffer = new InputBuffer(fileBytes);

						// Load the map
						this.skills.add(i, new Skill(fileBuffer));
					}
				} catch (IOException e) {
					MessageLogger.getInstance().log(e);
				}
			}
		}
		
		// Loading maps
		File mapFolder = new File(ServerConfiguration.serverPath, "maps");
		for (i = 0 ; i <= ServerConfiguration.getInstance().maxMaps ; i++)
		{
			File mapFile = new File(mapFolder.getPath(), "map"+i+".aoc");

			if (mapFile.length() > Integer.MAX_VALUE)
			{
				MessageLogger.getInstance().log("Erreur : Fichier "+mapFile.getAbsolutePath()+" trop grand.");
			}
			else
			{
				// Create a byte array
				byte fileBytes[] = new byte[(int)mapFile.length()];
				
				try {
					if (mapFile.exists())
					{
						FileInputStream fileStream = (new FileInputStream(mapFile));

						// Sock the file content in a buffer
						fileStream.read(fileBytes, 0, fileBytes.length);
						
						// Read the file buffer
						InputBuffer fileBuffer = new InputBuffer(fileBytes);

						// Load the map
						Map newMap = new Map(i, TypeTools.byteArrayToMD5(fileBytes), fileBuffer);
						this.environment.put(i, newMap);
					}
				} catch (IOException e) {
					MessageLogger.getInstance().log(e);
				}
			}
		}

		ArrayList<Short> allDreams = new ArrayList<Short>();
		for (Dream currentDream : this.dreams)
		{
			for (short currentMap : currentDream.maps)
			{
				allDreams.add(currentMap);
			}
		}
		
		TreeMap<Short, Map> environmentCopy = (TreeMap<Short, Map>)this.environment.clone();
		Set<Short> noDreamMaps = environmentCopy.keySet();
		noDreamMaps.removeAll(allDreams);
		
		for (short currentNoDreamMap : noDreamMaps)
		{
			this.getMap(currentNoDreamMap).newInstance();
		}
		
		// Loading areas
		File areaFolder = new File(ServerConfiguration.serverPath, "areas");
		for (i = 0 ; i <= ServerConfiguration.getInstance().maxAreas ; i++)
		{
			File areaFile = new File(areaFolder.getPath(), "area"+i+".aoz");

			if (areaFile.length() > Integer.MAX_VALUE)
			{
				MessageLogger.getInstance().log("Erreur : Fichier "+areaFile.getAbsolutePath()+" trop grand.");
			}
			else
			{
				// Create a byte array
				byte fileBytes[] = new byte[(int)areaFile.length()];
				
				try {
					if (areaFile.exists())
					{
						FileInputStream fileStream = (new FileInputStream(areaFile));

						// Sock the file content in a buffer
						fileStream.read(fileBytes, 0, fileBytes.length);
						
						// Read the file buffer
						InputBuffer fileBuffer = new InputBuffer(fileBytes);

						// Load the map
						this.cartography.add(i, new Area((byte)i, fileBuffer));
					}
				} catch (IOException e) {
					MessageLogger.getInstance().log(e);
				}
			}
		}
		
		// Loading items
		File itemFolder = new File(ServerConfiguration.serverPath, "items");
		for (i = 0 ; i <= ServerConfiguration.getInstance().maxItems ; i++)
		{
			File itemFile = new File(itemFolder.getPath(), "item"+i+".aoo");

			if (itemFile.length() > Integer.MAX_VALUE)
			{
				MessageLogger.getInstance().log("Erreur : Fichier "+itemFile.getAbsolutePath()+" trop grand.");
			}
			else
			{
				// Create a byte array
				byte fileBytes[] = new byte[(int)itemFile.length()];
				
				try {
					if (itemFile.exists())
					{
						FileInputStream fileStream = (new FileInputStream(itemFile));

						// Sock the file content in a buffer
						fileStream.read(fileBytes, 0, fileBytes.length);
						
						// Read the file buffer
						InputBuffer fileBuffer = new InputBuffer(fileBytes);

						// Load the items
						this.items.add(i, new Item(fileBuffer));
					}
				} catch (IOException e) {
					MessageLogger.getInstance().log(e);
				}
			}
		}
		
		// Loading crafts
		File craftFolder = new File(ServerConfiguration.serverPath, "crafts");
		for (i = 0 ; i <= ServerConfiguration.getInstance().maxCrafts ; i++)
		{
			File craftFile = new File(craftFolder.getPath(), "Craft"+i+".aop");

			if (craftFile.length() > Integer.MAX_VALUE)
			{
				MessageLogger.getInstance().log("Erreur : Fichier "+craftFile.getAbsolutePath()+" trop grand.");
			}
			else
			{
				// Create a byte array
				byte fileBytes[] = new byte[(int)craftFile.length()];
				
				try {
					if (craftFile.exists())
					{
						FileInputStream fileStream = (new FileInputStream(craftFile));

						// Sock the file content in a buffer
						fileStream.read(fileBytes, 0, fileBytes.length);
						
						// Read the file buffer
						InputBuffer fileBuffer = new InputBuffer(fileBytes);

						// Load the crafts
						this.crafts.add(i, new Craft(i, fileBuffer));
					}
				} catch (IOException e) {
					MessageLogger.getInstance().log(e);
				}
			}
		}
	
		// Launch weather timer
		ScheduledExecutorService weatherTimer = Executors.newSingleThreadScheduledExecutor();
		Runnable changeWeather = new TalkingRunnable(new Runnable() {
			@Override
			public void run() {
				synchronized(World.this.cartography)
				{
					for (Area currentArea : World.this.cartography)
					{
						currentArea.askGodForWeather();
					}
					
					Transmission.sendToAllPlayer(World.getInstance().getAreaWeatherPacket());
					//Transmission.sendWeatherToAll();
				}
			}
		});
		weatherTimer.scheduleAtFixedRate(changeWeather, 0, 5, TimeUnit.SECONDS);
		
		// Launch time timer
		ScheduledExecutorService timeTimer = Executors.newSingleThreadScheduledExecutor();
		Runnable changeTime = new TalkingRunnable(new Runnable() {
			@Override
			public void run() {
				
				World.this.timeLock.lock();
				
				World.this.time = !World.this.time;
				//Transmission.sendTimeToAll();
				Transmission.sendToAllPlayer(World.getInstance().getTimePacket());
				
				World.this.timeLock.unlock();
			}
		});
		timeTimer.scheduleAtFixedRate(changeTime, 0, 10, TimeUnit.SECONDS);
	}
	
	public synchronized final static World getInstance()
	{
		if (INSTANCE == null)
		{
			INSTANCE = new World();
		}
		return INSTANCE;
	}
	
	public Map getMap(int mapNum)
	{
		return this.environment.get((short)mapNum);
	}
	
	public boolean isDreamMap(int mapNum)
	{
		boolean rval = false;
		
		Iterator<Dream> iteDream = this.dreams.iterator();
		Dream currentDream;
		while (iteDream.hasNext() && !rval)
		{
			currentDream = iteDream.next();
			
			for (short currentMap : currentDream.maps)
			{
				if (currentMap == mapNum)
				{
					rval = true;
				}
			}
		}
		
		return rval;
	}
	
	public OutputBuffer getAreaWeatherPacket()
	{
		OutputBuffer packet = new OutputBuffer("SAreaWeather");
		
		packet.writeByte((byte)this.cartography.size());
		
		for (Area currentArea : this.cartography)
		{
			packet.writeByte(currentArea.getId());
			packet.writeByte(currentArea.getCurrentWeather().getCode());
		}
		
		return packet;
	}
	
	public OutputBuffer getTimePacket()
	{
		OutputBuffer packet = new OutputBuffer("STime");
		
		packet.writeBoolean(this.time);
		
		return packet;
	}

	/*public HashMap<Short, Map> getMaps()
	{
		return this.environment;
	}*/
}
