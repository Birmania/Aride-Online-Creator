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

import Communications.InputBuffer;
import Miscs.MessageLogger;
import Npc.NpcType;

public class Faun {
	private static Faun INSTANCE = null;
	
	public NpcType faun[];
	
	private Faun()
	{
		this.faun = new NpcType[ServerConfiguration.getInstance().maxNpcs];
		
		// Loading npc types
		File npcFolder = new File(ServerConfiguration.serverPath, "npcs");
		for (short i = 0 ; i <= ServerConfiguration.getInstance().maxNpcs ; i++)
		{
			File npcFile = new File(npcFolder.getPath(), "npc"+i+".aon");

			if (npcFile.length() > Integer.MAX_VALUE)
			{
				MessageLogger.getInstance().log("Erreur : Fichier "+npcFile.getAbsolutePath()+" trop grand.");
			}
			else
			{
				// Create a byte array
				byte fileBytes[] = new byte[(int)npcFile.length()];
				
				try {
					if (npcFile.exists())
					{
						FileInputStream fileStream = (new FileInputStream(npcFile));
						
						// Sock the file content in a buffer
						fileStream.read(fileBytes, 0, fileBytes.length);
						
						// Read the file buffer
						InputBuffer fileBuffer = new InputBuffer(fileBytes);
						
						// Load the map
						this.faun[i] = new NpcType(i, fileBuffer);
					}
				} catch (IOException e) {
					MessageLogger.getInstance().log(e);
				}
			}
		}
	}
	
	public synchronized final static Faun getInstance()
	{
		if (INSTANCE == null)
		{
			INSTANCE = new Faun();
		}
		return INSTANCE;
	}
}
