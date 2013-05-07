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

package Communications;

import java.lang.reflect.Array;
import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;

import Annotations.DeserializeIgnore;
import Annotations.StringSize;
import Miscs.MessageLogger;


public class BinaryFile {
	
	public void deserialize(InputBuffer buffer)
	{
		boolean mustProcessField;
		
		Field fields[] = this.getClass().getDeclaredFields();
		for (Field f : fields)
		{	
			mustProcessField = true;
			
			// Regarder si on doit le traiter
			if (f.getName().toString().compareTo("this$0") == 0) // Variable de classe externe (Outer)
			{
				mustProcessField = false;
			}
			
			try
			{	
				if (f.getAnnotations()[0].annotationType().equals(DeserializeIgnore.class))
				{
					mustProcessField = false;
				}
			} catch (ArrayIndexOutOfBoundsException e) {
				// No annotation, do nothing
			}
			
			if (mustProcessField)
			{
				if (f.getType().isArray())
				{
					short dimension = buffer.readShort();
					if (dimension > 0)
					{
						if (f.getType().equals(String[].class))
						{
							switch(dimension)
							{
								case 1:
								{
									String tab[] = new String[(int)buffer.readLong()];
									for (int i = 0 ; i < tab.length ; i++)
									{
										tab[i] = buffer.readString();
									}
									try {
										f.set(this, tab);
									} catch (IllegalArgumentException
											| IllegalAccessException e) {
										MessageLogger.getInstance().log(e);
									}
									break;
								}
								default:
									MessageLogger.getInstance().log("Trop de dimension pour le String. A implanté. Dimension : "+dimension);
							}
						}
						else if (f.getType().equals(byte[].class))
						{
							switch(dimension)
							{
								case 1:
								{
									byte tab[] = new byte[(int)buffer.readLong()];
									for (int i = 0 ; i < tab.length ; i++)
									{
										tab[i] = buffer.readByte();
									}
									try {
										f.set(this, tab);
									} catch (IllegalArgumentException
											| IllegalAccessException e) {
										MessageLogger.getInstance().log(e);
									}
									break;
								}
								default:
								{
									MessageLogger.getInstance().log("Trop de dimension pour les bytes. A implanté. Dimension : "+dimension);
								}
							}
						}
						else if (f.getType().equals(int[].class))
						{
							switch(dimension)
							{
								case 1:
								{
									int tab[] = new int[(int)buffer.readLong()];
									for (int i = 0 ; i < tab.length ; i++)
									{
										tab[i] = buffer.readInt();
									}
									try {
										f.set(this, tab);
									} catch (IllegalArgumentException
											| IllegalAccessException e) {
										MessageLogger.getInstance().log(e);
									}
									break;
								}
								default:
								{
									MessageLogger.getInstance().log("Trop de dimension pour les ints. A implanter. Dimension : "+dimension);
								}
							}
						}
						else if (f.getType().equals(short[].class))
						{
							switch(dimension)
							{
								case 1:
								{
									short tab[] = new short[(int)buffer.readLong()];
									for (int i = 0 ; i < tab.length ; i++)
									{
										tab[i] = buffer.readShort();
									}
									try {
										f.set(this, tab);
									} catch (IllegalArgumentException
											| IllegalAccessException e) {
										MessageLogger.getInstance().log(e);
									}
									break;
								}
								default:
								{
									MessageLogger.getInstance().log("Trop de dimension pour les ints. A implanter. Dimension : "+dimension);
								}
							}
						}
						else // Personnal object
						{
							switch(dimension)
							{
								case 1:
								{
									Object tab = Array.newInstance(f.getType().getComponentType(), (int)buffer.readLong());
									
							
									Class <?> type = f.getType().getComponentType();
									Constructor<?> ctor = type.getConstructors()[0];
									
									for (int i = 0 ; i < Array.getLength(tab) ; i++)
									{
								        Object val;
										try {
											if (type.getEnclosingClass() != null)
											{
												val = ctor.newInstance(this);
											}
											else
											{
												val = ctor.newInstance();
											}

											Array.set(tab, i, val);

											((BinaryFile)Array.get(tab, i)).deserialize(buffer);
										} catch (InstantiationException
												| IllegalAccessException
												| IllegalArgumentException
												| InvocationTargetException e) {
											MessageLogger.getInstance().log(e);
										}
									}
									try {
										f.set(this, tab);
									} catch (IllegalArgumentException
											| IllegalAccessException e) {
										MessageLogger.getInstance().log(e);
									}

									break;
								}
								case 2:
								{
									int secondDimension = (int)buffer.readLong();
									int firstDimension = (int)buffer.readLong();
									Object tab = Array.newInstance(f.getType().getComponentType().getComponentType(), firstDimension, secondDimension);
									
									Class <?> type = f.getType().getComponentType().getComponentType();
									Constructor<?> ctor = type.getConstructors()[0];
									
									for (int i = 0 ; i < Array.getLength(tab) ; i++)
									{
										Object tab2 = Array.newInstance(f.getType().getComponentType().getComponentType(), Array.getLength(Array.get(tab, i)));
										
										Array.set(tab, i , tab2);
									}
									
									for(int i = 0 ; i < Array.getLength(Array.get(tab, 0)) ; i++) // Nombre d'élément dans les deuxième dimension
									{
										for (int j = 0 ; j < Array.getLength(tab) ; j++) // Parcourir chaque tableau en lisant à chaque fois une valeur dans le buffer
										{
											Object val;
											try {
												if (type.getEnclosingClass() != null)
												{
													val = ctor.newInstance(this);
												}
												else
												{
													val = ctor.newInstance();
												}
												((BinaryFile)val).deserialize(buffer);
												Array.set(Array.get(tab, j), i, val);
											} catch (InstantiationException
													| IllegalAccessException
													| IllegalArgumentException
													| InvocationTargetException e) {
												MessageLogger.getInstance().log(e);
											}
										}
									}
									
									try {
										f.set(this, tab);
									} catch (IllegalArgumentException
											| IllegalAccessException e) {
										MessageLogger.getInstance().log(e);
									}
									break;
								}
								default:
								{
									MessageLogger.getInstance().log("Trop de dimension pour les objets. A implanté. Dimension : "+dimension);
								}
							}
						}
					}
				}
				else
				{
					if (f.getType().equals(String.class))
					{
						try {
							try {
								f.set(this, buffer.readString(((StringSize)f.getAnnotations()[0]).size()));
							} catch (IllegalArgumentException
									| IllegalAccessException e) {
								MessageLogger.getInstance().log(e);
							}
						} catch (ArrayIndexOutOfBoundsException e) {
							try {
								f.set(this, buffer.readString());
							} catch (IllegalArgumentException
									| IllegalAccessException e1) {
								MessageLogger.getInstance().log(e1);
								//e1.printStackTrace();
							}
						}
					}
					else if (f.getType().equals(int.class))
					{
						try {
							f.set(this, buffer.readInt());
						} catch (IllegalArgumentException | IllegalAccessException e) {
							MessageLogger.getInstance().log(e);
						}
					}
					else if (f.getType().equals(short.class))
					{
						try {
							f.set(this, buffer.readShort());
						} catch (IllegalArgumentException | IllegalAccessException e) {
							MessageLogger.getInstance().log(e);
						}
					}
					else if (f.getType().equals(byte.class))
					{
						try {
							f.set(this, buffer.readByte());
						} catch (IllegalArgumentException | IllegalAccessException e) {
							MessageLogger.getInstance().log(e);
						}
					}
					else if (f.getType().equals(boolean.class))
					{
						try {
							f.set(this, buffer.readBoolean());
						} catch (IllegalArgumentException | IllegalAccessException e) {
							MessageLogger.getInstance().log(e);
						}
					}
					else if (f.getType().equals(float.class))
					{
						try {
							f.set(this, buffer.readFloat());
						} catch (IllegalArgumentException | IllegalAccessException e) {
							MessageLogger.getInstance().log(e);
						}
					}
					else // Personnal object
					{
						try {
							Object obj = f.getType().newInstance();
							((BinaryFile)obj).deserialize(buffer);
						} catch (InstantiationException | IllegalAccessException e2) {
							MessageLogger.getInstance().log(e2);
							//e2.printStackTrace();
						}
					}
				}
			}
		}
	}
}
