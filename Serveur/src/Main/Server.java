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

import java.io.IOException;
import java.lang.management.ManagementFactory;
import java.lang.management.ThreadInfo;
import java.lang.management.ThreadMXBean;
import java.net.ServerSocket;
import java.util.ArrayList;
import java.util.Arrays;

import Communications.Transmission;
import Miscs.MessageLogger;
import Miscs.TypeTools;



public class Server {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		// Pour garder une précision de 1 milliseconde sous Windows on doit avoir ce thread en continu
		new Thread()
        {
 
            {
                this.setDaemon(true);
                this.start();
            }
 
            public void run()
            {
                while (true)
                {
                    try
                    {
                        Thread.sleep(Integer.MAX_VALUE);
                    } catch (InterruptedException ex)
                    {
                    }
                }
            }
        };
        
		new Thread()
        {
 
            {
                this.setDaemon(true);
                this.start();
            }
 
            public void run()
            {
                while (true)
                {
                	ThreadMXBean bean = ManagementFactory.getThreadMXBean();
                	long[] threadIds = bean.findDeadlockedThreads(); // Returns null if no threads are deadlocked.

                	if (threadIds != null) {
                	    ThreadInfo[] infos = bean.getThreadInfo(threadIds, true, true);

                	    ArrayList<Object> report = new  ArrayList<Object>();
                	    report.add("<<<<<<<<<<<<<<<< Deadlock detected ! >>>>>>>>>>>>>>>>");
                	    for (ThreadInfo info : infos) {
                	        StackTraceElement[] stack = info.getStackTrace();
                	        // Log or store stack trace information.
                	        report.addAll(Arrays.asList(stack));
                	        report.add(System.getProperty("line.separator"));
                	    }
                	    
                	    MessageLogger.getInstance().log(TypeTools.join(report.toArray(), System.getProperty("line.separator")));
                	    
                	    //TODO : Write in a file
                	}
                	try {
						Thread.sleep(3000);
					} catch (InterruptedException e) {
						// TODO Auto-generated catch block
						MessageLogger.getInstance().log(e);
					}
                }
            }
        };
		
		try {
			ServerSocket worldSocket = new ServerSocket(ServerConfiguration.getInstance().port); // ouverture d'un socket serveur sur port
			
			// L'ordre est important car le monde se crée à l'aide de la faune puis on spawn dans le monde
			Faun.getInstance();
			World.getInstance();
			
			// Initialiser le routeur
			Transmission.getInstance();
			
			// Initialiser le serveur configuration (va créer la connexion avec le serveur de données)
			ServerConfiguration.getInstance();
			
			Runtime.getRuntime().addShutdownHook(new Thread(Population.getInstance().new SaveAllPlayers(), "Aride-Shutdown-Thread"));
			
			while (true) // Attente en boucle de connexion (bloquant sur ss.accept)
			{
				Population.getInstance().addPlayer(worldSocket.accept()); // un client se connecte, un nouveau thread client est lancé
			}
		} catch (IOException e) {
			MessageLogger.getInstance().log(e);
		}

	}
}
