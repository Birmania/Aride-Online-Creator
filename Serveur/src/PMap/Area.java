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

import Annotations.DeserializeIgnore;
import Communications.BinaryFile;
import Communications.InputBuffer;
import Enumerations.WeatherTypes;

public class Area extends BinaryFile {
	@DeserializeIgnore
	private byte id;
	@DeserializeIgnore
	private WeatherTypes currentWeather;

	public String name;
	public float sandStormingFrequency;
	public float snowingFrequency;
	public float rainingFrequency;
	public float thunderingFrequency;
	
	public Area(byte id, InputBuffer buffer)
	{
		this.id = id;
		
		this.deserialize(buffer);
		
		this.askGodForWeather();
	}
	
	public void askGodForWeather()
	{
		double choice = Math.random();
		
		if (choice < this.sandStormingFrequency)
		{
			this.currentWeather = WeatherTypes.SandStorming;
		}
		else if (choice < (this.sandStormingFrequency+this.snowingFrequency))
		{
			this.currentWeather = WeatherTypes.Snowing;
		}
		else if (choice < (this.sandStormingFrequency+this.snowingFrequency+this.rainingFrequency))
		{
			this.currentWeather = WeatherTypes.Raining;
		}
		else if (choice < (this.sandStormingFrequency+this.snowingFrequency+this.rainingFrequency+this.thunderingFrequency))
		{
			this.currentWeather = WeatherTypes.Thundering;
		}
		else
		{
			this.currentWeather = WeatherTypes.None;
		}
	}
	
	public WeatherTypes getCurrentWeather()
	{
		return this.currentWeather;
	}
	
	public byte getId()
	{
		return this.id;
	}
}
