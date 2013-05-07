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

package PMap;

import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.ScheduledFuture;
import java.util.concurrent.TimeUnit;

import Enumerations.Directions;
import Miscs.TalkingRunnable;

public abstract class MapMovable extends MapElement {
	
	private Directions dir;
	
	private ScheduledExecutorService movementTimer;
	private Runnable movement;
	private ScheduledFuture<?> movementScheduler;
	public byte speed;
	public boolean moving;
	
	public MapMovable(Map.MapInstance map, byte x, byte y, Directions dir)
	{
		super(map, x, y);
		this.setDir(dir);

		this.movementTimer = Executors.newSingleThreadScheduledExecutor();
		this.movement = new TalkingRunnable(new Runnable() {

			@Override
			public void run() {
				MapMovable.this.move();
			}
			
		});
		this.speed = 0;
	}
	
	public void launchMovementTimer()
	{
		this.moving = true;
		this.movementScheduler = this.movementTimer.scheduleAtFixedRate(this.movement, 0, 1000/this.speed, TimeUnit.MILLISECONDS);
	}
	
	public void stopMovementTimer()
	{
		this.movementScheduler.cancel(false);
		this.moving = false;
	}
	
	public abstract void move();
	
	public void setDir(Directions dir)
	{
		this.dir = dir;
	}
	
	public Directions getDir()
	{
		return this.dir;
	}
	
}
