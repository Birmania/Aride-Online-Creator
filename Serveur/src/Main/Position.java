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

public class Position implements Comparable<Position>{
	
	private byte x;
	private byte y;
	
	public Position(byte x, byte y) {
		this.setX(x);
		this.setY(y);
	}
	
	public Position(Position position) {
		this.setX(position.getX());
		this.setY(position.getY());
	}
	
	public void setX(byte x)
	{
		this.x = x;
	}
	
	public byte getX()
	{
		return this.x;
	}
	
	public void setY(byte y)
	{
		this.y = y;
	}
	
	public byte getY()
	{
		return this.y;
	}
	
	public int distance(Position position)
	{
		return Math.abs(this.getX() - position.getX())+Math.abs(this.getY() - position.getY());
	}
	
	@Override
	public boolean equals(Object o)
	{
		boolean rval = false;
		
		if (o instanceof Position)
		{
			rval = (this.getX() == ((Position)o).getX()) && (this.getY() == ((Position)o).getY());
		}
		
		return rval;
	}

	@Override
	public int compareTo(Position o) {
		int rval = 0;
		
		if (!this.equals(o))
		{
			if (this.getY() > o.getY() || ((this.getY() == o.getY()) && (this.getX() > o.getX())))
			{
				rval = 1;
			}
			else
			{
				rval = -1;
			}
		}
		
		return rval;
	}

	// TODO : Eventuellement a supprimer en fin de dev
	@Override
	public String toString() {
		return "Position [x=" + x + ", y=" + y + "]";
	}
}
