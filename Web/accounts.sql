SET @@autocommit = 0; 

START TRANSACTION;
SET FOREIGN_KEY_CHECKS=0;

DROP TABLE IF EXISTS ItemSlot;
DROP TABLE IF EXISTS PlayerItem;
DROP TABLE IF EXISTS CraftSlot;
DROP TABLE IF EXISTS SkillSlot;
DROP TABLE IF EXISTS Player;

DROP PROCEDURE IF EXISTS create_player;
DROP PROCEDURE IF EXISTS save_player;
DROP PROCEDURE IF EXISTS set_item;
DROP PROCEDURE IF EXISTS add_skill;
DROP PROCEDURE IF EXISTS remove_skill;
DROP PROCEDURE IF EXISTS add_craft;
DROP PROCEDURE IF EXISTS remove_craft;

DROP VIEW IF EXISTS Account;
DROP VIEW IF EXISTS AccountItem;
DROP VIEW IF EXISTS AccountSkill;
DROP VIEW IF EXISTS AccountCraft;

DROP TRIGGER IF EXISTS t_ins_of_i_accountitem;

REVOKE ALL ON *.* FROM front_server;

CREATE TABLE CraftSlot
	(craftId SMALLINT UNSIGNED NOT NULL,
	playerId MEDIUMINT(8) UNSIGNED NOT NULL,
	INDEX crafts_of_player(playerId),
	PRIMARY KEY (craftId, playerId),
	FOREIGN KEY (playerId) REFERENCES Player(playerId) ON DELETE CASCADE
	) ENGINE=InnoDB;

CREATE TABLE SkillSlot
	(skillId SMALLINT UNSIGNED NOT NULL,
	playerId MEDIUMINT(8) UNSIGNED NOT NULL,
	INDEX skills_of_player(playerId),
	PRIMARY KEY (skillId, playerId),
	FOREIGN KEY (playerId) REFERENCES Player(playerId) ON DELETE CASCADE
	) ENGINE=InnoDB;

CREATE TABLE ItemSlot
	(idItemSlot INTEGER UNSIGNED AUTO_INCREMENT,
	itemId SMALLINT UNSIGNED NOT NULL,
	itemVal SMALLINT UNSIGNED NOT NULL,
	itemDur SMALLINT UNSIGNED NOT NULL,
	PRIMARY KEY (idItemSlot)
	) ENGINE=InnoDB;

CREATE TABLE PlayerItem
	(idItemSlot INTEGER UNSIGNED NOT NULL,
	playerId MEDIUMINT(8) UNSIGNED NOT NULL,
	numSlot TINYINT UNSIGNED NOT NULL,
	PRIMARY KEY (playerId, numSlot),
	FOREIGN KEY (idItemSlot) REFERENCES ItemSlot(idItemSlot) ON DELETE CASCADE,
	FOREIGN KEY (playerId) REFERENCES Player(playerId) ON DELETE CASCADE,
	UNIQUE (idItemSlot)
	) ENGINE=InnoDB;
	
CREATE TABLE Player
	(playerId MEDIUMINT(8) UNSIGNED NOT NULL,
	playerSex BOOLEAN NOT NULL,
	playerSprite SMALLINT UNSIGNED NOT NULL,
	playerLevel SMALLINT UNSIGNED NOT NULL,
	playerExp INTEGER UNSIGNED NOT NULL,
	playerLife INTEGER UNSIGNED NOT NULL,
	playerStamina INTEGER UNSIGNED NOT NULL,
	playerSleep INTEGER UNSIGNED NOT NULL,
	playerSTR SMALLINT UNSIGNED NOT NULL,
	playerDEF SMALLINT UNSIGNED NOT NULL,
	playerDEX SMALLINT UNSIGNED NOT NULL,
	playerSCI SMALLINT UNSIGNED NOT NULL,
	playerLANG SMALLINT UNSIGNED NOT NULL,
	playerFREE SMALLINT UNSIGNED NOT NULL,
	armorId INTEGER UNSIGNED DEFAULT NULL,
	weaponId INTEGER UNSIGNED DEFAULT NULL,
	helmetId INTEGER UNSIGNED DEFAULT NULL,
	shieldId INTEGER UNSIGNED DEFAULT NULL,
	petId SMALLINT UNSIGNED DEFAULT NULL,
	petLife INTEGER UNSIGNED DEFAULT NULL,
	playerMap SMALLINT UNSIGNED NOT NULL,
	playerX TINYINT UNSIGNED NOT NULL,
	playerY TINYINT UNSIGNED NOT NULL,
	PRIMARY KEY (playerId),
	FOREIGN KEY (armorId) REFERENCES ItemSlot(idItemSlot) ON DELETE SET NULL,
	FOREIGN KEY (weaponId) REFERENCES ItemSlot(idItemSlot) ON DELETE SET NULL,
	FOREIGN KEY (helmetId) REFERENCES ItemSlot(idItemSlot) ON DELETE SET NULL,
	FOREIGN KEY (shieldId) REFERENCES ItemSlot(idItemSlot) ON DELETE SET NULL,
	FOREIGN KEY (playerId) REFERENCES xxxxxxxxxx_users(user_id) ON DELETE CASCADE
	) ENGINE=InnoDB;
	
SET FOREIGN_KEY_CHECKS=1;

CREATE VIEW Account AS
	SELECT Plr.playerSex, Plr.playerSprite, Plr.playerLevel, Plr.playerLife, Plr.playerStamina, Plr.playerSleep, Plr.playerExp, Plr.playerSTR, Plr.playerDEF, Plr.playerDEX, Plr.playerSCI, Plr.playerLANG, Plr.playerFREE, Plr.petId, Plr.petLife, Plr.playerMap, Plr.playerX, Plr.playerY,
	Usr.username As "playerName", Usr.user_email As "playerEmail", Usr.user_password As "playerPassword",
	IS1.itemId As "playerArmorId", IS1.itemDur As "playerArmorDur", IS2.itemId As "playerWeaponId", IS2.itemDur As "playerWeaponDur", IS3.itemId As "playerHelmetId", IS3.itemDur As "playerHelmetDur", IS4.itemId As "playerShieldId", IS4.itemDur As "playerShieldDur"
	FROM Player Plr LEFT JOIN ItemSlot IS1 ON Plr.armorId = IS1.idItemSlot LEFT JOIN ItemSlot IS2 ON Plr.weaponId = IS2.idItemSlot LEFT JOIN ItemSlot IS3 ON Plr.helmetId = IS3.idItemSlot LEFT JOIN ItemSlot IS4 ON Plr.shieldId = IS4.idItemSlot, xxxxxxxxxx_users Usr
	WHERE Usr.user_id = Plr.playerId;

CREATE VIEW AccountItem As
	SELECT ISlot.itemId, ISlot.itemVal, ISlot.itemDur, PItem.numSlot, Usr.username As "playerName"
	FROM ItemSlot ISlot, PlayerItem PItem, Player Plr, xxxxxxxxxx_users Usr
	WHERE ISlot.idItemSlot = PItem.idItemSlot AND PItem.playerId = Plr.playerId AND Plr.playerId = Usr.user_id;
	
CREATE VIEW AccountSkill As
	SELECT SSlot.*, Usr.username As "playerName"
	FROM SkillSlot SSlot, Player Plr, xxxxxxxxxx_users Usr
	WHERE SSlot.playerId = Plr.playerId AND Plr.playerId = Usr.user_id;
	
CREATE VIEW AccountCraft As
	SELECT CSlot.*, Usr.username As "playerName"
	FROM CraftSlot CSlot, Player Plr, xxxxxxxxxx_users Usr
	WHERE CSlot.playerId = Plr.playerId AND Plr.playerId = Usr.user_id;
	
delimiter $$
CREATE PROCEDURE create_player(IN p_playerName VARCHAR(255), IN p_playerSex BOOLEAN, IN p_playerSprite SMALLINT UNSIGNED, IN p_playerMap SMALLINT UNSIGNED, IN p_playerX TINYINT UNSIGNED, IN p_playerY TINYINT UNSIGNED)
BEGIN
	DECLARE var_playerId MEDIUMINT(8) UNSIGNED;
	SELECT user_id INTO var_playerId FROM xxxxxxxxxx_users WHERE username = p_playerName;
	
	INSERT INTO Player(playerId, playerSex, playerSprite, playerLevel, playerExp, playerLife, playerStamina, playerSleep, playerSTR, playerDEF, playerDEX, playerSCI, playerLANG, playerFREE, playerMap, playerX, playerY) VALUES (var_playerId, p_playerSex, p_playerSprite, 1, 0, 200, 200, 200, 0, 0, 0, 0, 0, 10, p_playerMap, p_playerX, p_playerY);

	CALL add_skill(p_playerName, 0);
	CALL add_skill(p_playerName, 1);
	CALL add_skill(p_playerName, 2);
END$$

CREATE PROCEDURE save_player(IN p_playerName VARCHAR(255), IN p_playerMap SMALLINT UNSIGNED, IN p_playerX TINYINT UNSIGNED, IN p_playerY TINYINT UNSIGNED, IN p_playerLife INTEGER UNSIGNED, IN p_playerStamina INTEGER UNSIGNED, IN p_playerSleep INTEGER UNSIGNED, IN p_playerExp INTEGER UNSIGNED, IN p_playerLevel SMALLINT UNSIGNED, IN p_playerSTR SMALLINT UNSIGNED, IN p_playerDEF SMALLINT UNSIGNED, IN p_playerDEX SMALLINT UNSIGNED, IN p_playerSCI SMALLINT UNSIGNED, IN p_playerLANG SMALLINT UNSIGNED, IN p_playerFREE SMALLINT UNSIGNED, IN p_armorId SMALLINT UNSIGNED, IN p_armorDur SMALLINT UNSIGNED, IN p_weaponId SMALLINT UNSIGNED, IN p_weaponDur SMALLINT, IN p_helmetId SMALLINT UNSIGNED, IN p_helmetDur SMALLINT UNSIGNED, IN p_shieldId SMALLINT UNSIGNED, IN p_shieldDur SMALLINT UNSIGNED, IN p_petId SMALLINT UNSIGNED, IN p_petLife INTEGER UNSIGNED)
BEGIN
	DECLARE var_playerId MEDIUMINT(8) UNSIGNED;
	DECLARE var_armorId INTEGER UNSIGNED;
	DECLARE var_weaponId INTEGER UNSIGNED;
	DECLARE var_helmetId INTEGER UNSIGNED;
	DECLARE var_shieldId INTEGER UNSIGNED;

	
	SELECT user_id INTO var_playerId FROM xxxxxxxxxx_users WHERE username = p_playerName;
	
	UPDATE Account
		SET playerMap=p_playerMap, playerX=p_playerX, playerY=p_playerY, playerLife=p_playerLife, playerStamina=p_playerStamina, playerSleep=p_playerSleep, playerExp=p_playerExp, playerLevel=p_playerLevel, playerSTR=p_playerSTR, playerDEF=p_playerDEF, playerDEX=p_playerDEX, playerSCI=p_playerSCI, playerLANG=p_playerLANG, playerFREE=p_playerFREE, petId=p_petId, petLife=p_petLife
		WHERE playerName = p_playerName;
	SELECT Plr.armorId, Plr.weaponId, Plr.helmetId, Plr.shieldId INTO var_armorId, var_weaponId, var_helmetId, var_shieldId
	FROM Player Plr
	WHERE Plr.playerId = var_playerId;
	
	IF var_armorId IS NULL THEN
		IF p_armorId IS NOT NULL THEN
			INSERT INTO ItemSlot(itemId, itemVal, itemDur) VALUES(p_armorId, 1, p_armorDur);
			UPDATE Player
			SET armorId = LAST_INSERT_ID()
			WHERE playerId = var_playerId;
		END IF;
	ELSE
		IF p_armorId IS NULL THEN
			DELETE FROM ItemSlot
			WHERE idItemSlot = var_armorId;
		ELSE
			UPDATE Account
			SET playerArmorId = p_armorId, playerArmorDur = p_armorDur
			WHERE playerName = p_playerName;
		END IF;
	END IF;
	
	IF var_weaponId IS NULL THEN
		IF p_weaponId IS NOT NULL THEN
			INSERT INTO ItemSlot(itemId, itemVal, itemDur) VALUES(p_weaponId, 1, p_weaponDur);
			UPDATE Player
			SET weaponId = LAST_INSERT_ID()
			WHERE playerId = var_playerId;
		END IF;
	ELSE
		IF p_weaponId IS NULL THEN
			DELETE FROM ItemSlot
			WHERE idItemSlot = var_weaponId;
		ELSE
			UPDATE Account
			SET playerWeaponId = p_weaponId, playerWeaponDur = p_weaponDur
			WHERE playerName = p_playerName;
		END IF;
	END IF;
	
	IF var_helmetId IS NULL THEN
		IF p_helmetId IS NOT NULL THEN
			INSERT INTO ItemSlot(itemId, itemVal, itemDur) VALUES(p_helmetId, 1, p_helmetDur);
			UPDATE Player
			SET helmetId = LAST_INSERT_ID()
			WHERE playerId = var_playerId;
		END IF;
	ELSE
		IF p_helmetId IS NULL THEN
			DELETE FROM ItemSlot
			WHERE idItemSlot = var_helmetId;
		ELSE
			UPDATE Account
			SET playerHelmetId = p_helmetId, playerHelmetDur = p_helmetDur
			WHERE playerName = p_playerName;
		END IF;
	END IF;
	
	IF var_shieldId IS NULL THEN
		IF p_shieldId IS NOT NULL THEN
			INSERT INTO ItemSlot(itemId, itemVal, itemDur) VALUES(p_shieldId, 1, p_shieldDur);
			UPDATE Player
			SET shieldId = LAST_INSERT_ID()
			WHERE playerId = var_playerId;
		END IF;
	ELSE
		IF p_shieldId IS NULL THEN
			DELETE FROM ItemSlot
			WHERE idItemSlot = var_shieldId;
		ELSE
			UPDATE Account
			SET playerShieldId = p_shieldId, playerShieldDur = p_shieldDur
			WHERE playerName = p_playerName;
		END IF;
	END IF;
END$$

CREATE PROCEDURE set_item(IN p_playerName VARCHAR(255), IN p_numSlot TINYINT UNSIGNED, IN p_itemId SMALLINT UNSIGNED, IN p_itemVal SMALLINT UNSIGNED, IN p_itemDur SMALLINT UNSIGNED)
BEGIN
	DECLARE var_playerId MEDIUMINT(8) UNSIGNED;
	DECLARE var_idItemSlot INTEGER UNSIGNED;
	DECLARE var_itemId SMALLINT UNSIGNED;
	DECLARE var_itemVal SMALLINT UNSIGNED;
	DECLARE var_itemDur SMALLINT UNSIGNED;
	
	SELECT user_id INTO var_playerId FROM xxxxxxxxxx_users WHERE username = p_playerName;
	
	SELECT idItemSlot INTO var_idItemSlot FROM PlayerItem WHERE playerId=var_playerId AND numSlot=p_numSlot;

	IF var_idItemSlot THEN
		IF p_itemId IS NOT NULL THEN
			SELECT itemId, itemVal, itemDur INTO var_itemId, var_itemVal, var_itemDur FROM ItemSlot WHERE idItemSlot = var_idItemSlot;
			IF var_itemId <> p_itemid OR var_itemVal <> p_itemVal OR var_itemDur <> p_itemDur THEN
				UPDATE ItemSlot
				SET itemId = p_itemId, itemVal = p_itemVal, itemDur = p_itemDur
				WHERE idItemSlot = var_idItemSlot;
			END IF;
		ELSE
			DELETE FROM ItemSlot
			WHERE idItemSlot = var_idItemSlot;
		END IF;
	ELSE
		IF p_itemId IS NOT NULL THEN
			INSERT INTO ItemSlot(itemId, itemVal, itemDur) VALUES (p_itemId, p_itemVal, p_itemDur);
			INSERT INTO PlayerItem(playerId, numSlot, idItemSlot) VALUES (var_playerId, p_numSlot, LAST_INSERT_ID());
		END IF;
	END IF;
END$$

CREATE PROCEDURE add_skill(IN p_playerName VARCHAR(255), IN p_skillId SMALLINT UNSIGNED)
BEGIN
	DECLARE var_playerId MEDIUMINT(8) UNSIGNED;
	SELECT user_id INTO var_playerId FROM xxxxxxxxxx_users WHERE username = p_playerName;

	INSERT INTO SkillSlot(playerId, skillId) VALUES(var_playerId, p_skillId);
END$$

CREATE PROCEDURE remove_skill(IN p_playerName VARCHAR(255), IN p_skillId SMALLINT UNSIGNED)
BEGIN
	DECLARE var_playerId MEDIUMINT(8) UNSIGNED;
	SELECT user_id INTO var_playerId FROM xxxxxxxxxx_users WHERE username = p_playerName;

	DELETE 
	FROM SkillSlot
	WHERE playerId = var_playerId AND skillId = p_skillId;
END$$

CREATE PROCEDURE add_craft(IN p_playerName VARCHAR(255), IN p_craftId SMALLINT UNSIGNED)
BEGIN
	DECLARE var_playerId MEDIUMINT(8) UNSIGNED;
	SELECT user_id INTO var_playerId FROM xxxxxxxxxx_users WHERE username = p_playerName;

	INSERT INTO CraftSlot(playerId, craftId) VALUES(var_playerId, p_craftId);
END$$

CREATE PROCEDURE remove_craft(IN p_playerName VARCHAR(255), IN p_craftId SMALLINT UNSIGNED)
BEGIN
	DECLARE var_playerId MEDIUMINT(8) UNSIGNED;
	SELECT user_id INTO var_playerId FROM xxxxxxxxxx_users WHERE username = p_playerName;

	DELETE 
	FROM CraftSlot
	WHERE playerId = var_playerId AND craftId = p_craftId;
END$$
	
delimiter ;
	
GRANT SELECT ON Account TO front_server;
GRANT SELECT ON AccountItem TO front_server;
GRANT SELECT ON AccountSkill TO front_server;
GRANT SELECT ON AccountCraft TO front_server;
GRANT EXECUTE ON PROCEDURE save_player TO front_server;
GRANT EXECUTE ON PROCEDURE set_item TO front_server;
GRANT EXECUTE ON PROCEDURE add_skill TO front_server;
GRANT EXECUTE ON PROCEDURE remove_skill TO front_server;
GRANT EXECUTE ON PROCEDURE add_craft TO front_server;
GRANT EXECUTE ON PROCEDURE remove_craft TO front_server;

CALL create_player('Birmania', True, 0, 0, 0, 0);

CALL set_item('Birmania', 0, 0, 15, 1);
CALL set_item('Birmania', 1, 1, 1, 1);
CALL set_item('Birmania', 2, 2, 1, 1);
CALL set_item('Birmania', 3, 3, 1, 1);
CALL set_item('Birmania', 4, 6, 50, 1);


COMMIT;