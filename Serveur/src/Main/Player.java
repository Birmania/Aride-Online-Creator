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

import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.Semaphore;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.locks.Lock;
import java.util.concurrent.locks.ReadWriteLock;
import java.util.concurrent.locks.ReentrantLock;
import java.util.concurrent.locks.ReentrantReadWriteLock;

import Communications.OutputBuffer;
import Communications.Transmission;
import Enumerations.ItemTypes;
import Enumerations.MapElementTypes;
import Exceptions.NoNpcException;
import Exceptions.NoPlayerException;
import Exceptions.NoValidSpawnPositionException;
import Interfaces.IFighter;
import Interfaces.IKillable;
import Interfaces.IRecognizable;
import Miscs.MessageLogger;
import Miscs.TalkingRunnable;
import Miscs.TypeTools;
import PMap.Map;
import PMap.Map.MapInstance;
import PMap.MapWalkable;

public class Player extends MapWalkable implements IFighter, IKillable, IRecognizable{	
	// attributes
	// general informations
	public ClientThread client;
	
	private String name;
	public boolean sex; // False : Men, True : Women
	public short sprite;
	public short level;
	public short access; // 0 : classic player, 1 : admin, 2 : super-admin
	public String guild;
	
	// bars
	private int life;
	private int maxLife;
	private int stamina;
	private int maxStamina;
	private int sleep;
	private int maxSleep;
	public int exp;
	public float expModifier;
	
	public Lock fightLock;
	
	public long nextMovement;
	public long nextAttack;
	public long attackSpeed;
	
	// Statistics
	public short strength;
	public short strengthBonus;
	public short defense;
	public short defenseBonus;
	public short dexterity;
	public short dexterityBonus;
	public short science;
	public short scienceBonus;
	public short language;
	public short languageBonus;
	
	public short freePoints;
	
	public Lock statisticLock;
	
	// Equipement
	private ItemSlot armorSlot;
	private ItemSlot weaponSlot;
	private ItemSlot helmetSlot;
	private ItemSlot shieldSlot;
	//public ItemSlot petSlot;
	
	// Position
	public Pet pet;
	
	// Inventory
	public ItemSlot inventory[] = new ItemSlot[ServerConfiguration.getInstance().maxPlayerItems+1];
	
	public Lock inventoryLock;
	
	// Effects like potions
	public ArrayList<Item> effectItems = new ArrayList<Item>();
	
	// Skills
	public List<Short> skills = new ArrayList<Short>(ServerConfiguration.getInstance().maxPlayerSkills);
	//private List<Short> loadedSkills = new ArrayList<Short>(ServerConfiguration.getInstance().maxPlayerSkills);
	private List<Short> loadedSkills;
	public Lock skillsLock;
	
	// Quests
	public short quests[] = new short[ServerConfiguration.getInstance().maxPlayerQuests];
	
	// Crafts
	//public short crafts[] = new short[ServerConfiguration.getInstance().maxPlayerCrafts];
	public List<Short> crafts = new ArrayList<Short>();
	private List<Short> loadedCrafts;
	
	// Party
	public Party party;
	
	// TODO : To Remove
	//public long test;
	
	private ScheduledExecutorService tirednessTimer;
	
	//private boolean inWarp;
	public Semaphore warpLock;
	public ReadWriteLock partyLock;
	public Lock petLock;
	
	public byte movementController; // 0 : Prédit, 1 : Forcé
	
	public Dream.DreamInstance dreamInstance;
	public Map.MapInstance mapBeforeDream;
	public Position positionBeforeDream;
	
	public Player(ClientThread client, Map.MapInstance map, byte x, byte y, ResultSet account) throws SQLException
	{
		super(map, x, y);
		
		this.client = client;
		
		// First thing to do is to lock the warp
		this.warpLock = new Semaphore(0);

		this.partyLock = new ReentrantReadWriteLock(true);
		this.petLock = new ReentrantLock();
		this.skillsLock = new ReentrantLock();
	
		// General informations
		this.name = account.getString("playerName");
		this.sex = account.getBoolean("playerSex");
		this.sprite = account.getShort("playerSprite");
		this.level = account.getShort("playerLevel");
		// TODO : access ou pas ?
		this.access = 0;
		// TODO : Guild
		this.guild = null;
		
		// Load bars
		this.life = account.getInt("playerLife");
		this.maxLife = ServerConfiguration.getInstance().maxPlayerLife;
		this.stamina = account.getInt("playerStamina");
		this.maxStamina = ServerConfiguration.getInstance().maxPlayerStamina;
		this.sleep = account.getInt("playerSleep");
		this.maxSleep = ServerConfiguration.getInstance().maxPlayerSleep;
		this.exp = account.getInt("playerExp");
		this.expModifier = 0;
		
		this.fightLock = new ReentrantLock();
		
		this.nextAttack = 0;
		this.attackSpeed = ServerConfiguration.getInstance().baseAttackSpeed;
		
		// Load statistics
		this.strength = account.getShort("playerSTR");
		this.strengthBonus = 0;
		this.defense = account.getShort("playerDEF");
		this.defenseBonus = 0;
		this.dexterity = account.getShort("playerDEX");
		this.dexterityBonus = 0;
		this.science = account.getShort("playerSCI");
		this.scienceBonus = 0;
		this.language = account.getShort("playerLANG");
		this.languageBonus = 0;
		
		this.freePoints = account.getShort("playerFREE");
		
		this.statisticLock = new ReentrantLock();
		
		this.inventoryLock = new ReentrantLock();
		
		// Equipement. Each value refer to the player inventory
		short armorId = account.getShort("playerArmorId");
		short armorVal = -1;
		short armorDur = -1;
		if (account.wasNull())
		{
			armorId = -1;
		}
		else
		{
			armorVal = 1;
			armorDur = account.getShort("playerArmorDur");
		}
		//this.armorSlot = new ItemSlot(armorId, armorVal, armorDur);
		this.setArmorSlot(new ItemSlot(armorId, armorVal, armorDur));
		
		short weaponId = account.getShort("playerWeaponId");
		short weaponVal = -1;
		short weaponDur = -1;
		if (account.wasNull())
		{
			weaponId = -1;
		}
		else
		{
			weaponVal = 1;
			weaponDur = account.getShort("playerWeaponDur");
		}
		//this.weaponSlot = new ItemSlot(weaponId, weaponVal, weaponDur);
		this.setWeaponSlot(new ItemSlot(weaponId, weaponVal, weaponDur));
		
		short helmetId = account.getShort("playerHelmetId");
		short helmetVal = -1;
		short helmetDur = -1;
		if (account.wasNull())
		{
			helmetId = -1;
		}
		else
		{
			helmetVal = 1;
			helmetDur = account.getShort("playerHelmetDur");
		}
		//this.helmetSlot = new ItemSlot(helmetId, helmetVal, helmetDur);
		this.setHelmetSlot(new ItemSlot(helmetId, helmetVal, helmetDur));
		
		short shieldId = account.getShort("playerShieldId");
		short shieldVal = -1;
		short shieldDur = -1;
		if (account.wasNull())
		{
			shieldId = -1;
		}
		else
		{
			shieldVal = 1;
			shieldDur = account.getShort("playerShieldDur");
		}
		//this.shieldSlot = new ItemSlot(shieldId, shieldVal, shieldDur);
		this.setShieldSlot(new ItemSlot(shieldId, shieldVal, shieldDur));
		/*this.weaponSlot = new ItemSlot(Short.parseShort(character.fetch("WeaponId")), Short.parseShort(character.fetch("WeaponVal")), Short.parseShort(character.fetch("WeaponDur")));
		this.helmetSlot = new ItemSlot(Short.parseShort(character.fetch("HelmetId")), Short.parseShort(character.fetch("HelmetVal")), Short.parseShort(character.fetch("HelmetDur")));
		this.shieldSlot = new ItemSlot(Short.parseShort(character.fetch("ShieldId")), Short.parseShort(character.fetch("ShieldVal")), Short.parseShort(character.fetch("ShieldDur")));*/
		
		// Position
		// Inutile de faire le petLock.lock() ici. On va le faire pour la forme
		this.petLock.lock();
		short petId = account.getShort("petId");
		if (account.wasNull())
		{
			this.pet = null;
		}
		else
		{
			this.pet = new Pet(this.getId(), map, x, y, Faun.getInstance().faun[petId]);
		}
		this.petLock.unlock();
			
		// Inventory
		/*try {
			Thread.sleep(10000);
		} catch (InterruptedException e) {
			// TODO Auto-generated catch block
			MessageLogger.getInstance().log(e);
		}*/
		Connection con = ServerConfiguration.getInstance().getConnection();
		int i = 0;
		for (i = 0 ; i <= ServerConfiguration.getInstance().maxPlayerItems ; i++)
		{
			this.inventory[i] = new ItemSlot((short)-1, (short)-1, (short)-1);
		}
		ResultSet items = ServerConfiguration.getInstance().sendSelectQuery(con, "SELECT * FROM AccountItem WHERE playerName='"+this.getName()+"';");
		if (items != null)
		{
			while (items.next())
			{
				this.inventory[items.getByte("numSlot")] = new ItemSlot(items.getShort("itemId"), items.getShort("itemVal"), items.getShort("itemDur"));
				//i++;
			}
		}
		
		ServerConfiguration.getInstance().releaseConnection(con);
		/*for (i = i ; i <= ServerConfiguration.getInstance().maxPlayerItems ; i++)
		{
			this.inventory[i] = new ItemSlot((short)-1, (short)-1, (short)-1);
		}*/
		
		/*int i;
		for (i = 0 ; i <= ServerConfiguration.getInstance().maxPlayerItems ; i++)
		{
			this.inventory[i] = new ItemSlot(Short.parseShort(character.fetch("InvItemId"+i)), Short.parseShort(character.fetch("InvItemVal"+i)), Short.parseShort(character.fetch("InvItemDur"+i)));
		}*/
		/*int i;
		for (i = 0 ; i <= ServerConfiguration.getInstance().maxPlayerItems ; i++)
		{
			this.inventory[i] = new ItemSlot((short)-1, (short)-1, (short)-1);
		}*/
		
		// skills
		con = ServerConfiguration.getInstance().getConnection();
		ResultSet skills = ServerConfiguration.getInstance().sendSelectQuery(con, "SELECT * FROM AccountSkill WHERE playerName='"+this.getName()+"';");
		i = 0;
		if (skills != null)
		{
			while (skills.next())
			{
				this.skills.add(i, skills.getShort("skillId"));
				i++;
			}
		}
		ServerConfiguration.getInstance().releaseConnection(con);
		
		for (i = i ; i < ServerConfiguration.getInstance().maxPlayerSkills ; i++)
		{
			this.skills.add(i, (short)-1);
		}
		
		this.loadedSkills = new ArrayList<Short>(this.skills);
		
		// crafts
		con = ServerConfiguration.getInstance().getConnection();
		ResultSet crafts = ServerConfiguration.getInstance().sendSelectQuery(con, "SELECT * FROM AccountCraft WHERE playerName='"+this.getName()+"';");
		i = 0;
		if (crafts != null)
		{
			while (crafts.next())
			{
				this.crafts.add(i, crafts.getShort("craftId"));
				i++;
			}
		}
		ServerConfiguration.getInstance().releaseConnection(con);
		
		/*for (i = i ; i < ServerConfiguration.getInstance().maxPlayerCrafts ; i++)
		{
			this.crafts.add(i, (short)-1);
		}*/
		
		this.loadedCrafts = new ArrayList<Short>(this.crafts);
		
		// Skills
		/*for (i = 0 ; i < ServerConfiguration.getInstance().maxPlayerSkills ; i++)
		{
			this.skills[i] = Short.parseShort(character.fetch("Skill"+i));
		}
		
		// Quests
		for (i = 1 ; i <= ServerConfiguration.getInstance().maxPlayerQuests ; i++)
		{
			this.quests[i-1] = Short.parseShort(character.fetch("Quest"+i));
		}
		
		// Crafts
		for (i = 1 ; i <= ServerConfiguration.getInstance().maxPlayerCrafts ; i++)
		{
			this.crafts[i-1] = Short.parseShort(character.fetch("Craft"+i));
		}*/
		
		this.dreamInstance = null;
		this.mapBeforeDream = null;
		this.positionBeforeDream = null;
		
		this.movementController = 0;
	}
	
	public void prepareToSave()
	{
		this.prepareToFight();
		this.inventoryLock.lock();
		this.statisticLock.lock();
		this.skillsLock.lock();
		this.petLock.lock();
	}
	
	public void releaseFromSave()
	{
		this.petLock.unlock();
		this.skillsLock.unlock();
		this.statisticLock.unlock();
		this.inventoryLock.unlock();
		this.escapeFromFight();
	}
	
	public String savePlayer()
	{
		long a = System.currentTimeMillis();
		String rval = "";
		/*this.prepareToFight();
		this.inventoryLock.lock();
		this.statisticLock.lock();
		this.skillsLock.lock();
		this.petLock.lock();*/
		this.prepareToSave();
		
		int playerMap = this.getMapInstance().getMap().id;
		byte playerX = this.getX();
		byte playerY = this.getY();
		if (this.dreamInstance != null)
		{
			playerMap = this.mapBeforeDream.getMap().id;
			playerX =  this.positionBeforeDream.getX();
			playerY = this.positionBeforeDream.getY();
		}
		
		if (this.isDead())
		{
			this.life = this.maxLife/2;
		}
		
		Short petId = null;
		Integer petLife = null;
		
		if (this.pet != null)
		{
			petId = this.pet.type.id;
			petLife = this.pet.getLife();
		}
		//this.petLock.unlock();
		
		// TODO : peut être travaillé directement dans ItemSlot avec des objets

		/*synchronized(this.inventory)
		{*/
		// Save equipment
		
		Short armorId = this.armorSlot.getItemId();
		Short armorDur = this.armorSlot.getItemDur();
		if (armorId == -1)
		{
			armorId = null;
			armorDur = null;
		}
		Short weaponId = this.weaponSlot.getItemId();
		Short weaponDur = this.weaponSlot.getItemDur();
		if (weaponId == -1)
		{
			weaponId = null;
			weaponDur = null;
		}
		Short helmetId = this.helmetSlot.getItemId();
		Short helmetDur = this.helmetSlot.getItemDur();
		if (helmetId == -1)
		{
			helmetId = null;
			helmetDur = null;
		}
		Short shieldId = this.shieldSlot.getItemId();
		Short shieldDur = this.shieldSlot.getItemDur();
		if (shieldId == -1)
		{
			shieldId = null;
			shieldDur = null;
		}
	
		
		//ServerConfiguration.getInstance().sendRequest("CALL save_player('" +
		rval += "CALL save_player('" +
				this.getName()+"',"+
				playerMap+","+
				playerX+","+
				playerY+","+
				this.life+","+
				this.stamina+","+
				this.sleep+","+
				this.exp+","+
				this.level+","+
				this.strength+","+
				this.defense+","+
				this.dexterity+","+
				this.science+","+
				this.language+","+
				this.freePoints+","+
				armorId+","+
				armorDur+","+
				weaponId+","+
				weaponDur+","+
				helmetId+","+
				helmetDur+","+
				shieldId+","+
				shieldDur+","+
				petId+","+
				petLife
				+ ");";

		// Save inventory
		int i;
		Short itemId;
		Short itemVal;
		Short itemDur;
		for (i = 0 ; i <= ServerConfiguration.getInstance().maxPlayerItems ; i++)
		{
			itemId = this.inventory[i].getItemId();
			itemVal = this.inventory[i].getItemVal();
			itemDur = this.inventory[i].getItemDur();
			if (itemId == -1)
			{
				itemId = null;
				itemVal = null;
				itemDur = null;
			}
			
			rval += "CALL set_item("+
					"'"+this.getName()+"',"+
					i+","+
					itemId+","+
					itemVal+","+
					itemDur+
					//");");
					");";
		}
		//}
		//this.inventoryLock.unlock();

		/*synchronized(this.inventory)
		{
			for (i = 0 ; i <= ServerConfiguration.getInstance().maxPlayerItems ; i++)
			{
				itemId = this.inventory[i].getItemId();
				itemVal = this.inventory[i].getItemVal();
				itemDur = this.inventory[i].getItemDur();
				if (itemId == -1)
				{
					itemId = null;
					itemVal = null;
					itemDur = null;
				}
				
				rval += "CALL set_item("+
						"'"+this.getName()+"',"+
						i+","+
						itemId+","+
						itemVal+","+
						itemDur+
						//");");
						");";
			}
		}*/
		
		// Skills
		/*synchronized(this.skills)
		{*/
		List<Short> skillsToRemove = TypeTools.substract(this.loadedSkills, this.skills);
		Iterator<Short> ite = skillsToRemove.iterator();
		while (ite.hasNext())
		{
			rval += "CALL remove_skill('"+this.getName()+"', "+ite.next()+");";
			//ServerConfiguration.getInstance().sendRequest("CALL remove_skill('"+this.getName()+"', "+ite.next()+");");
		}
		
		List<Short> skillsToAdd = TypeTools.substract(this.skills, this.loadedSkills);
		ite = skillsToAdd.iterator();
		while (ite.hasNext())
		{
			rval += "CALL add_skill('"+this.getName()+"', "+ite.next()+");";
			//ServerConfiguration.getInstance().sendRequest("CALL add_skill('"+this.getName()+"', "+ite.next()+");");
		}
		
		this.loadedSkills = new ArrayList<Short>(this.skills);
		//}
		
		// Crafts
		synchronized(this.crafts)
		{
			List<Short> craftsToRemove = TypeTools.substract(this.loadedCrafts, this.crafts);
			ite = craftsToRemove.iterator();
			while (ite.hasNext())
			{
				rval += "CALL remove_craft('"+this.getName()+"', "+ite.next()+");";
				//ServerConfiguration.getInstance().sendRequest("CALL remove_skill('"+this.getName()+"', "+ite.next()+");");
			}
			
			List<Short> craftsToAdd = TypeTools.substract(this.crafts, this.loadedCrafts);
			ite = craftsToAdd.iterator();
			while (ite.hasNext())
			{
				rval += "CALL add_craft('"+this.getName()+"', "+ite.next()+");";
				//ServerConfiguration.getInstance().sendRequest("CALL add_skill('"+this.getName()+"', "+ite.next()+");");
			}
			
			this.loadedCrafts = new ArrayList<Short>(this.crafts);
		}

		this.releaseFromSave();
		
		return rval;
	}
	
	public short getId()
	{
		return this.client.getId();
	}
	
	public ItemSlot getShieldSlot()
	{
		return this.shieldSlot;
	}
	
	public void setShieldSlot(ItemSlot newSlot)
	{
		if (this.shieldSlot != null && this.shieldSlot.getItemId() >= 0)
		{
			this.removeItemEffects(World.getInstance().items.get(this.shieldSlot.getItemId()));
		}
		this.shieldSlot = newSlot;
		if (newSlot.getItemId() >= 0)
		{
			this.applyItemEffects(World.getInstance().items.get(newSlot.getItemId()));
		}
	}
	
	public ItemSlot getHelmetSlot()
	{
		return this.helmetSlot;
	}
	
	public void setHelmetSlot(ItemSlot newSlot)
	{
		if (this.helmetSlot != null && this.helmetSlot.getItemId() >= 0)
		{
			this.removeItemEffects(World.getInstance().items.get(this.helmetSlot.getItemId()));
		}
		this.helmetSlot = newSlot;
		if (newSlot.getItemId() >= 0)
		{
			this.applyItemEffects(World.getInstance().items.get(newSlot.getItemId()));
		}
	}
	
	public ItemSlot getArmorSlot()
	{
		return this.armorSlot;
	}
	
	public void setArmorSlot(ItemSlot newSlot)
	{
		if (this.armorSlot != null && this.armorSlot.getItemId() >= 0)
		{
			this.removeItemEffects(World.getInstance().items.get(this.armorSlot.getItemId()));
		}
		this.armorSlot = newSlot;
		if (newSlot.getItemId() >= 0)
		{
			this.applyItemEffects(World.getInstance().items.get(newSlot.getItemId()));
		}
	}

	public ItemSlot getWeaponSlot()
	{
		return this.weaponSlot;
	}
	
	public void setWeaponSlot(ItemSlot newSlot)
	{
		if (this.weaponSlot != null && this.weaponSlot.getItemId() >= 0)
		{
			this.removeItemEffects(World.getInstance().items.get(this.weaponSlot.getItemId()));
		}
		this.weaponSlot = newSlot;
		if (newSlot.getItemId() >= 0)
		{
			this.applyItemEffects(World.getInstance().items.get(newSlot.getItemId()));
		}
	}
	
	/*public void setInWarp(boolean value)
	{
		this.inWarp = value;
	}
	
	public boolean isInWarp()
	{
		return this.inWarp;
	}*/
	
	public void startTiredness()
	{
		this.tirednessTimer = Executors.newSingleThreadScheduledExecutor();
		
		Runnable sleepChanger = new TalkingRunnable(new Runnable() {

			@Override
			public void run() {
				Player.this.sleepChange();
			}
			
		});
		
		this.tirednessTimer.scheduleAtFixedRate(sleepChanger, 1, 1, TimeUnit.SECONDS);
	}
	
	public void stopTiredness()
	{
		this.tirednessTimer.shutdownNow();
	}
	
	protected OutputBuffer getStopMovePacket()
	{	
		OutputBuffer packet = new OutputBuffer("SPlayerStopMove");
		
		packet.writeShort(this.getId());
		
		packet.writeByte(this.getX());

		packet.writeByte(this.getY());
		
		return packet;
	}
	
	public void sendStopMovePacketToMapBut()
	{
		Transmission.sendToMapInstanceBut(this.getMapInstance(), this, this.getStopMovePacket());
	}
	
	public void sendStopMovePacketToMap()
	{
		this.movementController = 1;
		Transmission.sendToMapInstance(this.getMapInstance(), this.getStopMovePacket());
	}
	
	public void sendStopMovePacketToPlayer()
	{
		this.movementController = 1;
		Transmission.sendToClient(this.client, this.getStopMovePacket());
	}
	
	public OutputBuffer getStartMovePacket()
	{
		OutputBuffer packet = new OutputBuffer("SPlayerStartMove");
		
		packet.writeShort(this.getId());
		packet.writeByte(this.getDir().getCode());
		//packet.writeByte((byte)(1000/this.movementTimer.getDelay()));
		packet.writeByte(this.speed);
		
		return packet;
	}
	
	public OutputBuffer getDirMovePacket()
	{
		OutputBuffer packet = new OutputBuffer("SPlayerDirMove");
		
		packet.writeShort(this.getId());
		
		packet.writeByte(this.getDir().getCode());
		 
		packet.writeByte(this.getX()); // X position
		packet.writeByte(this.getY()); // Y position
		
		return packet;
	}
	
	public OutputBuffer getDirPacket()
	{
		OutputBuffer packet = new OutputBuffer("SPlayerDir");
		
		packet.writeShort(this.getId());
		
		packet.writeByte(this.getDir().getCode());
		
		return packet;
	}
	
	public void writeEquipmentInPacket(OutputBuffer packet)
	{
		packet.writeShort(this.getArmorSlot().getItemId());
		packet.writeShort(this.getArmorSlot().getItemVal());
		packet.writeShort(this.getArmorSlot().getItemDur());
		packet.writeShort(this.getWeaponSlot().getItemId());
		packet.writeShort(this.getWeaponSlot().getItemVal());
		packet.writeShort(this.getWeaponSlot().getItemDur());
		packet.writeShort(this.getHelmetSlot().getItemId());
		packet.writeShort(this.getHelmetSlot().getItemVal());
		packet.writeShort(this.getHelmetSlot().getItemDur());
		packet.writeShort(this.getShieldSlot().getItemId());
		packet.writeShort(this.getShieldSlot().getItemVal());
		packet.writeShort(this.getShieldSlot().getItemDur());
	}
	
	public void writeLifeInPacket(OutputBuffer packet)
	{	
		// Write the life
		packet.writeInt(this.getMaxLife());
		packet.writeInt(this.getLife());
	}
	
	public void writeStaminaInPacket(OutputBuffer packet)
	{
		// Write the endurance
		packet.writeInt(ServerConfiguration.getInstance().maxPlayerStamina);
		packet.writeInt(this.stamina);
	}
	
	public void writeSleepInPacket(OutputBuffer packet)
	{
		// Write the sleep
		packet.writeInt(ServerConfiguration.getInstance().maxPlayerSleep);
		packet.writeInt(this.sleep);
	}
	
	public void writeExperienceInPacket(OutputBuffer packet)
	{
		// Write the experience
		packet.writeInt(this.exp);
	}
	
	public void enterGame()
	{
		this.enterMapInstanceSafe(this.getMapInstance(), this.getX(), this.getY());
		this.client.setInGame(true);
		
		Transmission.sendPlayerInventory(this);
		
		Transmission.sendPlayerStamina(this);

		Transmission.sendPlayerSleep(this);
		Transmission.sendPlayerNextLevel(this);
		Transmission.sendPlayerExperience(this);
		
		Transmission.sendPlayerLife(this);
		
		Transmission.sendPlayerSkills(this);
		
		Transmission.sendPlayerCrafts(this);
		
		Transmission.sendPlayerStatistics(this);
		
		// Send every name to the player and the player name to everyone
		Transmission.communicateNewPlayer(this);
		
		Transmission.sendWeatherToPlayer(this);
			
		Transmission.sendTimeToPlayer(this);
		
		this.startTiredness();
	}
	
	public void warp(Map.MapInstance map, byte x, byte y)
	{
		//try {
		//this.warpLock.acquire();

		//this.setInWarp(true);
		
		//this.movementTimer.stop();
		/*this.stopMovementTimer();
		
		this.petLock.lock();
		if (this.pet != null)
		{
			this.pet.prepareToFight();
			if (!this.pet.isDead())
			{
				this.pet.stopMovementTimer();
			}
			this.pet.escapeFromFight();
		}
		this.petLock.unlock();*/
		
		this.quitMap();
		
		this.enterMapInstanceSafe(map, x, y);
		
		// Doit être en dernier au cas où le joueur déconnecte pendant un changement de map de rêve vers l'extérieur
		/*if (!World.getInstance().isDreamMap(map.getMap().id) && this.dreamInstance != null)
		{
			this.dreamInstance = null;
			this.mapBeforeDream = null;
			this.positionBeforeDream = null;
		}*/
		//} catch (InterruptedException e) {
		//	MessageLogger.getInstance().log(e);
		//}
	}
	
	public void enterMapInstanceSafe(Map.MapInstance map, byte x, byte y)
	{
		// Set the player map and position
		this.setMapInstance(map);
		
		// Attention : Pas d'allocation de tile ici !!! ca sera fait dans enterMap
		// D'ailleurs on refait ce genre de manip dans enterMapInstance... => Mais c'est a cause du sleep
		this.getPosition().setX(x);
		this.getPosition().setY(y);
		//this.setXY(x, y);
		
		OutputBuffer packet = new OutputBuffer("SCheckForMap");
		
		packet.writeShort(map.getMap().id);
		packet.writeString(map.getMap().md5);
		
		//this.sendPacket(packet);
		Transmission.sendToClient(this.client, packet);
		
		// Doit être en dernier au cas où le joueur déconnecte pendant un changement de map de rêve vers l'extérieur
		if (!World.getInstance().isDreamMap(map.getMap().id) && this.dreamInstance != null)
		{
			this.dreamInstance = null;
			this.mapBeforeDream = null;
			this.positionBeforeDream = null;
		}
	}
	
	public void quitMap()
	{	
		if (this.moving)
		{
			this.stopMovementTimer();
		}
		
		this.petLock.lock();
		if (this.pet != null)
		{
			this.pet.prepareToFight();
			if (!this.pet.isDead())
			{
				this.getMapInstance().deleteObserver(this.pet);
				if (this.pet.moving)
				{
					this.pet.stopMovementTimer();
				}
				//this.pet.attackTimer.stop();
				this.pet.stopAttackTimer();
			}
			this.pet.escapeFromFight();
		}
		this.petLock.unlock();
		
		// Remove the player from the current map
		this.getMapInstance().removePlayer(this);
		
		OutputBuffer packet = new OutputBuffer("SQuitMap");
		packet.writeShort(this.getId());
		Transmission.sendToMapInstanceBut(this.getMapInstance(), this, packet);
	}

	
	public void enterMapInstance()
	{	
		// TODO : Locker la map instance pour ne pas avoir de changement apres avoir obtenu notre position de spawn
		Position spawnPosition = null;
		while (spawnPosition == null)
		{
			try {
				spawnPosition = this.getMapInstance().getClosestSpawnPosition(this);
			} catch (NoValidSpawnPositionException e) {
				// TODO Auto-generated catch block
				MessageLogger.getInstance().log(e);
				try {
					Thread.sleep(500);
				} catch (InterruptedException e1) {
					// TODO Auto-generated catch block
					MessageLogger.getInstance().log(e1);
					//e1.printStackTrace();
				}
			}
		}	
		
		
		// Add the player in the map
		
		this.getMapInstance().sendLivingNpcsTo(this); // Important de le faire avant le addPlayer car les NPC écoutent la map et peuvent envoyer des messages au client (donc avant que le client connaisse les npc)
		this.getMapInstance().sendItemsTo(this);

		// On fait un getPosition pour ne pas faire d'allocation de tile (sera fait par le addplayer)
		this.getPosition().setX(spawnPosition.getX());
		this.getPosition().setY(spawnPosition.getY());
		this.getMapInstance().addPlayer(this, this.getX(), this.getY());
		
		
		this.getMapInstance().sendLivingPlayersTo(this); // Doit être fait après le addPlayer	
		
		// Send player data to everyone on the map except himself
		//Transmission.sendToMapInstanceBut(this, this.getPositionPacket());
		OutputBuffer packet = new OutputBuffer("SPlayerPosition");
		packet.writeShort((short)1);
		// TODO : writeInPacket must write the pet position
		this.petLock.lock();
		this.writePositionInPacket(packet);
		Transmission.sendToMapInstanceBut(this.getMapInstance(), this, packet);
		
		if (this.pet != null)
		{
			this.pet.followMaster();
		}
		this.petLock.unlock();
	}
	
	public void writeStartInfosInPacket(OutputBuffer packet)
	{	
		packet.writeString(this.name);
		packet.writeShort(this.sprite);
		this.partyLock.readLock().lock();
		if (this.party != null)
		{
			packet.writeShort(this.party.getId());
		}
		else
		{
			packet.writeShort((byte)-1);
		}
		this.partyLock.readLock().unlock();
		
		if (this.pet != null)
		{
			this.pet.writeStartInfosInPacket(packet);
		}
		else
		{
			packet.writeShort((short)-1);
		}
	}
	
	public void writePositionInPacket(OutputBuffer packet)
	{
		// Player position
		packet.writeShort(this.getId());
		packet.writeShort(this.getX());
		packet.writeShort(this.getY());
		packet.writeByte(this.getDir().getCode());
		if (this.moving)
		{
			packet.writeByte(this.speed);
		}
		else
		{
			packet.writeByte((byte)0);
		}
		
		// Player's pet position
		if (this.pet != null)
		{
			this.pet.writePositionInPacket(packet);
		}
	}
	
	public byte moveOnPosition(byte x, byte y)
	{
		byte rval = 0;
		
		if (this.getX() != x || this.getY() != y) // Si il est déjà sur la position on ne fait rien car sinon il ne peut pas allouer sa propre case
		{
			if (this.getMapInstance().getTileAllocation(x, y).isTraversableBy(this))
			{
				this.setXY(x, y);
			}
			else
			{
				rval = 2;
				this.stopMovementTimer();
				this.sendStopMovePacketToMap();
			}
		}
		else
		{
			rval = 1;
		}

		return rval;
	}
	
	/*public byte moveOnPosition(byte x, byte y)
	{
		byte rval = 0;
		
		if (this.getX() != x || this.getY() != y) // Si il est déjà sur la position on ne fait rien car sinon il ne peut pas allouer sa propre case
		{
			MapElement atDestination = this.getMap().getTileAllocation(x, y).getMapWalkable();
			if (atDestination == null)
			{
				this.setXY(x, y);
			}
			else
			{
				if (atDestination instanceof MapWalkable)
				{
					rval = 2;
					switch(this.getDir()) { // Est-ce que le joueur fait façe à l'élément dérangeant
					case DIR_UP:
						if (((MapWalkable)atDestination).getDir() == Directions.DIR_DOWN)
						{
							rval = 3;
						}
						break;
					case DIR_DOWN:
						if (((MapWalkable)atDestination).getDir() == Directions.DIR_UP)
						{
							rval = 3;
						}
						break;
					case DIR_LEFT:
						if (((MapWalkable)atDestination).getDir() == Directions.DIR_RIGHT)
						{
							rval = 3;
						}
						break;
					case DIR_RIGHT:
						if (((MapWalkable)atDestination).getDir() == Directions.DIR_LEFT)
						{
							rval = 3;
						}
						break;
					}
				}
				else
				{
					rval = 3;
				}
				if (rval == 3)
				{
					//this.movementTimer.stop(); // Normalement devrait déjà être fait bien plus tôt mais au cas où...
					this.stopMovementTimer();
					this.sendStopMovePacketToMap();
				}
				else if (rval == 2)
				{
					this.setXY(x, y); // Se fera sans allocation de la case
				}
			}
		}
		else
		{
			rval = 1;
		}

		return rval;
	}*/
	
	public void tryToAttack(MapElementTypes targetType, short targetIndex) throws NoNpcException, NoPlayerException
	{
		IKillable target = null;
		switch(targetType)
		{
		case Npc:
			target = this.getMapInstance().getNpc(targetIndex);
			break;
		case Player:
			target = Population.getInstance().getPlayer(targetIndex).player;
			break;
		}
				
		this.attack(target);
		/*switch(targetType)
		{
		case Npc:
			try {
				this.attack(this.getMapInstance().getNpc(targetIndex));
			} catch (NoNpcException e) {
				MessageLogger.getInstance().log(e);
			}
			break;
		case Player:
			try {
				this.attack(Population.getInstance().getPlayer(targetIndex).player);
			} catch (NoPlayerException e) {
				MessageLogger.getInstance().log(e);
			}
			break;
		}*/
	}
	
	public void attack(IKillable target)
	{
		this.attack(target, this.getDamage());
	}
	
	public void attack(IKillable target, int damage)
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
			
			/*OutputBuffer packet = new OutputBuffer("SDamageDisplay");
			packet.writeByte(MapElementTypes.Player.getCode());
			packet.writeShort(this.getId());
			packet.writeByte(MapElementTypes.get((MapElement)target).getCode());
			packet.writeShort(target.getId());
			packet.writeInt(damage);
			this.sendPacket(packet);*/
			
			//packet.writeString("Tu as fait "+damage+" dégats à "+target.getName());
			//packet.writeShort(Colors.White.getCode());
			
			Transmission.sendDamageDisplay(this, target, damage);
			if (target.removeLife(damage)) // target killed
			{
				this.exp += target.getExpValue() + target.getExpValue()*this.expModifier;
				Transmission.sendPlayerExperience(this);
			}
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

	@Override
	public int getLife() {
		return this.life;
	}
	
	public int getMaxLife() {
		return this.maxLife;
	}
	
	public int getStamina() {
		return this.stamina;
	}
	
	public int getMaxStamina() {
		return this.maxStamina;
	}
	
	public int getSleep() {
		return this.sleep;
	}
	
	public int getMaxSleep() {
		return this.maxSleep;
	}

	@Override
	public int getDamage() {
		int rval = this.strength+this.strengthBonus;
		
		// Inventory must be locked
		if (this.getWeaponSlot().getItemId() >= 0)
		{
			Item weapon = World.getInstance().items.get(this.getWeaponSlot().getItemId());
			if (weapon.type == ItemTypes.ItemTypeMissile.getCode() || weapon.type == ItemTypes.ItemTypeThrowable.getCode())
			{
				rval = this.dexterity+this.dexterityBonus;
			}
		}
		
		if (rval < 0)
		{
			rval = 0;
		}
		
		return rval;
	}

	@Override
	public boolean removeLife(int life) {
		boolean rval = false;
		
		this.life -= life;
		if (this.life > 0)
		{
			// Do not send everytime
			//Transmission.sendPlayerLife(this);
		}
		else
		{
			rval = true;
			try {
				this.warpLock.acquire();
			} catch (InterruptedException e) {
				throw new Error();
			}
			this.quitMap();
			
			OutputBuffer packet = new OutputBuffer("SPlayerDead");
			Transmission.sendToClient(this.client, packet);
			
			ScheduledExecutorService respawnTimer = Executors.newSingleThreadScheduledExecutor();
			Runnable respawner = new TalkingRunnable(new Runnable() {

				@Override
				public void run() {
					Player.this.respawn();
				}
				
			});
			respawnTimer.schedule(respawner, 5, TimeUnit.SECONDS);
		}
		
		return rval;
	}
	
	public void addLife(int life)
	{
		this.life += life;
		
		if (this.life > this.maxLife)
		{
			this.life = this.maxLife;
		}
		
		//Transmission.sendPlayerLife(this);
	}
	
	public void addSleep(int value)
	{
		this.sleep += value;
		
		if (this.sleep > this.maxSleep)
		{
			this.sleep = this.maxSleep;
		}
	}
	
	public void respawn() {
		this.partyLock.readLock().lock();
		if (this.party != null)
		{
			this.party.membersLock.lock();
		}

		MapInstance ancientMap = this.getMapInstance();
		ancientMap.lock.lock();
		this.prepareToFight(); // Because if not done, the enterMapInstanceSafe could change the map of the player during a kill or a save (if client exiter is executed)
		
		////
		Map.MapInstance newMap = World.getInstance().getMap(0).getOriginInstance();
		byte newX = (byte)0;
		byte newY = (byte)0;
		if (this.dreamInstance != null)
		{
			newMap = this.mapBeforeDream;
			newX = this.positionBeforeDream.getX();
			newY = this.positionBeforeDream.getY();
			
			this.dreamInstance = null;
			this.mapBeforeDream = null;
			this.positionBeforeDream = null;
		}
		
		this.setMapInstance(newMap);
		
		// Attention : Pas d'allocation de tile ici !!! ca sera fait dans enterMap
		this.getPosition().setX(newX);
		this.getPosition().setY(newY);
		////
		
		this.life = this.maxLife/2;
		this.enterMapInstanceSafe(this.getMapInstance(), this.getX(), this.getY());
		//this.warp(newMap, newX, newY);
		Transmission.sendPlayerLife(this);
		this.escapeFromFight();
		ancientMap.lock.unlock();
		
		if (this.party != null)
		{
			this.party.membersLock.unlock();
		}
		this.partyLock.readLock().unlock();
	}

	@Override
	public String getName() {
		return this.name;
	}
	
	/*public void sendPacket(OutputBuffer packet) throws DoNotUseException
	{
		this.client.sendPacket(packet);
	}*/

	/*class test implements Runnable {

		@Override
		public void run() {
			byte[] te = new byte[10];
			InputBuffer buff = new InputBuffer(te);
			
			try {
				HandleData.getInstance().handle(ClientThread.this, 15, te);
			} catch (IOException e) {
				// TODO Auto-generated catch block
				MessageLogger.getInstance().log(e);
			}
		}
		
	}*/
	
	@Override
	public void move() {
		this.getMapInstance().lock.lock();
		this.prepareToFight();
		if (!this.isDead())
		{
			if (this.moving)
			{
				this.moveOneOnDir();
			}
		}
		this.escapeFromFight();
		this.getMapInstance().lock.unlock();
	}
	
	public int getInventoryPosition(int itemId)
	{
		int rval = -1;
		
		int i = 0;

		while (rval == -1 && i < this.inventory.length)
		{
			if (this.inventory[i].getItemId() == itemId)
			{
				rval = i;
			}
			i++;
		}
		
		return rval;
	}
	
	/*public boolean addItem(ItemSlot itemSlot)
	{
		boolean added = false;
		synchronized(this.inventory)
		{
			int i = 0;
			
			while (i < this.inventory.length && !added)
			{
				if (this.inventory[i].getItemId() == -1)
				{
					this.inventory[i].setItemId(itemSlot.getItemId());
					this.inventory[i].addItemVal(itemSlot.getItemVal());
					this.inventory[i].setItemDur(itemSlot.getItemDur());
					//this.inventory[i] = itemSlot;
					
					Transmission.sendPlayerInventorySlot(this, (byte)i);
					
					added = true;
				}
				i++;
			}
		}
		return added;
	}*/
	
	public boolean addItem(ItemSlot itemSlot)
	{
		byte slotToProcess = -1;
		this.inventoryLock.lock();
		/*synchronized(this.inventory)
		{*/
		/*int i = 0;
		
		while (i < this.inventory.length && !added)
		{
			if (this.inventory[i].getItemId() == -1)
			{
				this.inventory[i].setItemId(itemSlot.getItemId());
				this.inventory[i].addItemVal(itemSlot.getItemVal());
				this.inventory[i].setItemDur(itemSlot.getItemDur());
				//this.inventory[i] = itemSlot;
				
				Transmission.sendPlayerInventorySlot(this, (byte)i);
				
				added = true;
			}
			i++;
		}*/
		if (World.getInstance().items.get(itemSlot.getItemId()).empilable == 1)
		{
			byte i = 0;
			
			while (i < this.inventory.length && slotToProcess == -1)
			{
				if (this.inventory[i].getItemId() == itemSlot.getItemId())
				{
					slotToProcess = i;
					this.inventory[slotToProcess].addItemVal(itemSlot.getItemVal());
				}
				
				i++;
			}
		}
		
		if (slotToProcess == -1)
		{
			slotToProcess = this.getFreeItemSlot();
			if (slotToProcess >= 0)
			{
				this.inventory[slotToProcess].setItemId(itemSlot.getItemId());
				this.inventory[slotToProcess].addItemVal(itemSlot.getItemVal());
				this.inventory[slotToProcess].setItemDur(itemSlot.getItemDur());
			}
		}
		
		if (slotToProcess >= 0)
		{
			Transmission.sendPlayerInventorySlot(this, slotToProcess);
		}
		//}
		this.inventoryLock.unlock();
		return (slotToProcess >= 0);
	}
	
	public void applyItemEffects(Item item)
	{
		short lifeEffect = item.lifeEffect;
		if (lifeEffect > 0)
		{
			this.addLife(lifeEffect);
		}
		else if (lifeEffect < 0)
		{
			this.removeLife(lifeEffect);
		}
		
		short sleepEffect = item.sleepEffect;
		if (sleepEffect > 0)
		{
			this.addSleep(sleepEffect);
		}
		else if (sleepEffect < 0)
		{
			this.removeSleep(sleepEffect);
		}
		
		this.maxLife += item.addHP;
		if (this.getLife() > this.getMaxLife())
		{
			this.life = this.getMaxLife();
		}
		
		this.maxStamina += item.addSTP;
		if (this.getStamina() > this.getMaxStamina())
		{
			this.stamina = this.getMaxStamina();
		}
		
		this.maxSleep += item.addSLP;
		if (this.getSleep() > this.getMaxSleep())
		{
			this.sleep = this.getMaxSleep();
		}
		
		this.expModifier += item.addExp / 1000;
		
		this.statisticLock.lock();
		
		this.strengthBonus += item.addStr;
		this.defenseBonus += item.addDef;
		this.scienceBonus += item.addSci;
		this.dexterityBonus += item.addDex;
		this.languageBonus += item.addLang;
		
		this.statisticLock.unlock();
		
		this.attackSpeed += this.attackSpeed * (item.attackSpeed / 1000);
	}
	
	public void removeItemEffects(Item item)
	{
		this.maxLife -= item.addHP;
		if (this.getLife() > this.getMaxLife())
		{
			this.life = this.getMaxLife();
		}
		
		this.expModifier -= item.addExp / 1000;
		
		this.statisticLock.lock();
		
		this.strengthBonus -= item.addStr;
		this.defenseBonus -= item.addDef;
		this.scienceBonus -= item.addSci;
		this.dexterityBonus -= item.addDex;
		this.languageBonus -= item.addLang;
		
		this.statisticLock.unlock();
		
		this.attackSpeed -= this.attackSpeed * (item.attackSpeed / 1000);
	}
	
	private byte getFreeItemSlot()
	{
		byte rval = -1;
		this.inventoryLock.lock();
		/*synchronized(this.inventory)
		{*/
		byte i = 0;
		
		while (i < this.inventory.length && rval == -1)
		{
			if (this.inventory[i].getItemId() == -1)
			{
				rval = i;
			}
			i++;
		}
		//}
		this.inventoryLock.unlock();
		return rval;
	}
	
	public void deleteItem(byte inventorySlot, short num)
	{
		this.inventoryLock.lock();
		/*synchronized(this.inventory)
		{*/
		this.inventory[inventorySlot].reduceItemVal(num);
		Transmission.sendPlayerInventorySlot(this, inventorySlot);
		//}
		this.inventoryLock.unlock();
	}
	
	@Override
	public boolean isDead()
	{
		return this.getLife() <= 0 || !this.client.isInGame(); // Comparaison du thread à null dans le cas où le joueur déco (pour que les NPC arretent de l'attaquer)
	}
	
	@Override
	public boolean tryToPrepareToFight() { // Not used for the moment
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
	
	public void sleepChange()
	{
		/*if (ClientThread.Player.this.sleep > 0)
		{
			int newValue = ClientThread.Player.this.sleep - 1;
			if (newValue < 0)
			{
				newValue = 0;
			}
			ClientThread.Player.this.sleep = newValue;
			
			Transmission.sendPlayerSleep(ClientThread.Player.this);
			
		}*/
		if (this.sleep > 0)
		{
			this.removeSleep(1);
			Transmission.sendPlayerSleep(this);
		}
	}
	
	public void removeSleep(int value)
	{
		int newValue = this.sleep - value;
		if (newValue < 0)
		{
			newValue = 0;
		}
		Player.this.sleep = newValue;
	}

	@Override
	public int getExpValue() {
		return 0;
	}
	
	/*class SleepChanger implements Runnable {

		@Override
		public void run() {

		}
		
	}*/
}