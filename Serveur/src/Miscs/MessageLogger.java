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

package Miscs;


import java.io.File;
import java.io.IOException;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.util.logging.FileHandler;
import java.util.logging.Level;
import java.util.logging.Logger;

import Main.ServerConfiguration;


public class MessageLogger {
	private static MessageLogger INSTANCE = null;
	
	private Logger logger;
	
	private MessageLogger()
	{
		FileHandler file;
		try {
			file = new FileHandler(new File(ServerConfiguration.serverPath, "logs.xml").getPath());
			
			this.logger = Logger.getLogger("logger");
			this.logger.addHandler(file);
		} catch (SecurityException | IOException e) {
			MessageLogger.getInstance().log(e);
			System.exit(-1);
		}
	}
	
	public synchronized final static MessageLogger getInstance()
	{
		if (INSTANCE == null)
		{
			INSTANCE = new MessageLogger();
		}
		return INSTANCE;
	}
	
	public void log(String chaine) {
		this.logger.log(Level.WARNING, chaine);
	}
	
	public void log(Exception e) {
		StringWriter sw = new StringWriter();
		PrintWriter pw = new PrintWriter(sw);
		e.printStackTrace(pw);
		this.log(sw.toString());
	}
}
