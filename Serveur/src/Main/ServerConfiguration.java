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

import java.beans.PropertyVetoException;
import java.io.File;
import java.io.IOException;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;

import org.ini4j.Ini;

import Miscs.MessageLogger;

import com.mchange.v2.c3p0.ComboPooledDataSource;


public class ServerConfiguration {
	private static ServerConfiguration INSTANCE = null;
	
	public static final byte nameLength = 40;
	
	private ComboPooledDataSource dataServerLink;
	
	public static String serverPath = "data";
	
	public static String secCode1 = "kwehiehfojcvnvnsdinaoiwheoewyriusdyrflsdjncjkxzncisdughfusyfuapsipiuahfpaijnflkjnvjnuahguiryasbdlfkjblsahgfauygewuifaunfauf";
	public static String secCode2 = "lsisyshentwuegeguigdfjkldsnoksamdihuehfidsuhdushdsisjsyayejrioehdoisahdjlasndowijapdnaidhaioshnksfnifohaifhaoinfiwnfinsaihfas";
	public static String secCode3 = "taiugdapuigoihwbdpiaugsdcapvhvinbudhbpidusbnvduisysayaspiufhpijsanfioasnpuvnupashuasohdaiofhaosifnvnuvnuahiosaodiubasdi";
	public static String secCode4 = "98978465734619123425676749756722829121973794379467987945762347631462572792798792492416127957989742945642672";
	
	public int port;
	public int dataServerPort;
	public String webSite;
	public String acceptedClientVersion;
	
	public short maxPlayers;
	public short maxAreas;
	public short maxItems;
	public short maxCrafts;
	public short maxMaps;
	public short maxNpcs;
	public short maxDreams;
	public short maxSkills;
	
	public short maxPlayerItems;
	public short maxPlayerSkills;
	public short maxPlayerQuests;
	public short maxPlayerCrafts;
	public int maxPlayerLife;
	public int maxPlayerStamina;
	public int maxPlayerSleep;
	public short maxPartyMembers;
	
	public long baseAttackSpeed = 1000;
	
	public ScheduledExecutorService scheduledExecutor;
	
	private ServerConfiguration()
	{
		this.scheduledExecutor = Executors.newScheduledThreadPool(100);
		
		Ini conf;
		try {
			conf = new Ini(new File(serverPath, "Data.ini"));
			
			this.port = Integer.parseInt(conf.get("CONFIG").fetch("Port"))+1;
			
			this.webSite = conf.get("CONFIG").fetch("WebSite");
			this.dataServerPort = Integer.parseInt(conf.get("CONFIG").fetch("DataServerPort"));
			
			// Ouvrir la connexion avec la base de données joueurs
			try {
				this.dataServerLink = new ComboPooledDataSource();
				this.dataServerLink.setDriverClass("com.mysql.jdbc.Driver");
				this.dataServerLink.setJdbcUrl("jdbc:mysql://"+this.webSite+":"+this.dataServerPort+"/forumdev?allowMultiQueries=true");
				this.dataServerLink.setUser("front_server");
				this.dataServerLink.setPassword("front_server_password");
				this.dataServerLink.setMinPoolSize(5);
				this.dataServerLink.setMaxPoolSize(5);
				Connection allCon[] = new Connection[this.dataServerLink.getMinPoolSize()];
				// Initialisation des connexions
				for (int i = 0 ; i < allCon.length ; i++)
				{
					allCon[i] = this.getConnection();
					// Update query type is not very useful here...
					this.sendUpdateQuery(allCon[i], "SET SESSION wait_timeout=2147483;"); // set pseudo-unlimited timeout
				}
				for (int i = 0 ; i < allCon.length ; i++)
				{
					this.releaseConnection(allCon[i]);
				}
			} catch (PropertyVetoException e1) {
				// TODO Auto-generated catch block
				MessageLogger.getInstance().log(e1);
			} catch (SQLException e) {
				MessageLogger.getInstance().log(e);
				System.exit(1);
			}
			
			this.acceptedClientVersion = conf.get("CONFIG").fetch("AcceptedClientVersion");
			
			this.maxPlayers = Short.parseShort(conf.get("MAX").fetch("MAX_PLAYERS"));
			
			this.maxAreas = Short.parseShort(conf.get("MAX").fetch("MAX_AREAS"));
			
			this.maxItems = Short.parseShort(conf.get("MAX").fetch("MAX_ITEMS"));
			
			this.maxCrafts = Short.parseShort(conf.get("MAX").fetch("MAX_CRAFTS"));
			
			this.maxMaps = Short.parseShort(conf.get("MAX").fetch("MAX_MAPS"));
			
			this.maxNpcs = Short.parseShort(conf.get("MAX").fetch("MAX_NPCS"));
			
			this.maxDreams = Short.parseShort(conf.get("MAX").fetch("MAX_DREAMS"));
			
			this.maxSkills = Short.parseShort(conf.get("MAX").fetch("MAX_SKILLS"));
			
			this.maxPlayerItems = Short.parseShort(conf.get("MAX").fetch("MAX_PLAYER_ITEMS"));
			
			this.maxPlayerSkills = Short.parseShort(conf.get("MAX").fetch("MAX_PLAYER_SKILLS"));
			
			this.maxPlayerQuests = Short.parseShort(conf.get("MAX").fetch("MAX_PLAYER_QUESTS"));
			
			this.maxPlayerCrafts = Short.parseShort(conf.get("MAX").fetch("MAX_PLAYER_CRAFTS"));
			
			this.maxPlayerLife = Short.parseShort(conf.get("MAX").fetch("MAX_PLAYER_LIFE"));
			
			this.maxPlayerStamina = Short.parseShort(conf.get("MAX").fetch("MAX_PLAYER_STAMINA"));
			
			this.maxPlayerSleep = Short.parseShort(conf.get("MAX").fetch("MAX_PLAYER_SLEEP"));
			
			this.maxPartyMembers = Short.parseShort(conf.get("MAX").fetch("MAX_PARTY_MEMBERS"));
		} catch (IOException e) {
			MessageLogger.getInstance().log(e);
		}
	}
	
	public final static ServerConfiguration getInstance() // Devrait être synchronized
	{
		if (INSTANCE == null)
		{
			INSTANCE = new ServerConfiguration();
		}
		return INSTANCE;
	}
	
	public Connection getConnection() throws SQLException
	{	
		return (Connection)this.dataServerLink.getConnection();
	}
	
	public void releaseConnection(Connection con) throws SQLException
	{
		con.close();
	}
	
	public ResultSet sendSelectQuery(Connection con, String request) throws SQLException
	{
		ResultSet rval = null;
		Statement statement = (Statement) con.createStatement();
		rval = statement.executeQuery(request);
		return rval;
	}
	
	public void sendUpdateQuery(Connection con, String request) throws SQLException
	{
		Statement statement = (Statement) con.createStatement();
		try {
			request = "START TRANSACTION;" + request;
			request = request + "COMMIT;";
			statement.executeQuery(request);
		} catch (SQLException e) {
			statement.executeQuery("ROLLBACK;");
			throw e;
		}
	}
}
