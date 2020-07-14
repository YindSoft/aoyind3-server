-- MySQL dump 10.13  Distrib 5.7.17, for Win64 (x86_64)
--
-- Host: localhost    Database: aoyind3
-- ------------------------------------------------------
-- Server version	5.7.19-log

/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8 */;
/*!40103 SET @OLD_TIME_ZONE=@@TIME_ZONE */;
/*!40103 SET TIME_ZONE='+00:00' */;
/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;
/*!40111 SET @OLD_SQL_NOTES=@@SQL_NOTES, SQL_NOTES=0 */;

--
-- Table structure for table `clanes`
--

DROP TABLE IF EXISTS `clanes`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `clanes` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Founder` varchar(20) DEFAULT NULL,
  `GuildName` varchar(20) DEFAULT NULL,
  `Fecha` date DEFAULT NULL,
  `Antifaccion` bigint(20) DEFAULT '0',
  `Alineacion` tinyint(4) DEFAULT '0',
  `Desc` varchar(500) DEFAULT NULL,
  `GuildNews` varchar(500) DEFAULT NULL,
  `URL` varchar(80) DEFAULT NULL,
  `Leader` varchar(20) DEFAULT NULL,
  `Codex1` varchar(80) DEFAULT NULL,
  `Codex2` varchar(80) DEFAULT NULL,
  `Codex3` varchar(80) DEFAULT NULL,
  `Codex4` varchar(80) DEFAULT NULL,
  `Codex5` varchar(80) DEFAULT NULL,
  `Codex6` varchar(80) DEFAULT NULL,
  `Codex7` varchar(80) DEFAULT NULL,
  `Codex8` varchar(80) DEFAULT NULL,
  `CantMiembros` int(11) DEFAULT '0',
  `EleccionesFinalizan` date DEFAULT NULL,
  PRIMARY KEY (`Id`),
  UNIQUE KEY `Id` (`Id`)
) ENGINE=MyISAM AUTO_INCREMENT=13 DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `clanes_propuestas`
--

DROP TABLE IF EXISTS `clanes_propuestas`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `clanes_propuestas` (
  `Id` bigint(20) NOT NULL AUTO_INCREMENT,
  `IdClan` int(11) DEFAULT NULL,
  `IdClanTo` int(11) DEFAULT NULL,
  `Detalle` varchar(400) DEFAULT NULL,
  `Tipo` int(11) DEFAULT '0',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `Id` (`Id`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `clanes_relaciones`
--

DROP TABLE IF EXISTS `clanes_relaciones`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `clanes_relaciones` (
  `IdClan` int(11) NOT NULL,
  `IdClanTo` int(11) NOT NULL,
  `Relacion` int(11) DEFAULT '0',
  PRIMARY KEY (`IdClan`,`IdClanTo`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `clanes_solicitudes`
--

DROP TABLE IF EXISTS `clanes_solicitudes`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `clanes_solicitudes` (
  `Id` bigint(20) NOT NULL AUTO_INCREMENT,
  `IdClan` int(11) DEFAULT NULL,
  `Nombre` varchar(20) DEFAULT NULL,
  `Solicitud` varchar(400) DEFAULT NULL,
  PRIMARY KEY (`Id`),
  UNIQUE KEY `Id` (`Id`)
) ENGINE=MyISAM AUTO_INCREMENT=159 DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `clanes_votos`
--

DROP TABLE IF EXISTS `clanes_votos`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `clanes_votos` (
  `IdClan` int(11) NOT NULL,
  `Nombre` varchar(20) NOT NULL,
  `Voto` varchar(20) DEFAULT NULL,
  PRIMARY KEY (`IdClan`,`Nombre`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `conexiones`
--

DROP TABLE IF EXISTS `conexiones`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `conexiones` (
  `Id` bigint(20) NOT NULL AUTO_INCREMENT,
  `IdPj` bigint(20) NOT NULL,
  `IP` varchar(15) DEFAULT NULL,
  `Fecha` datetime DEFAULT NULL,
  PRIMARY KEY (`Id`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `cuentas`
--

DROP TABLE IF EXISTS `cuentas`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `cuentas` (
  `Id` bigint(20) NOT NULL AUTO_INCREMENT,
  `Nombre` varchar(20) NOT NULL,
  `Password` varchar(32) DEFAULT NULL,
  `Email` varchar(100) DEFAULT NULL,
  `NombreApellido` varchar(150) DEFAULT NULL,
  `Direccion` varchar(250) DEFAULT NULL,
  `Ciudad` varchar(40) DEFAULT NULL,
  `Pais` varchar(2) DEFAULT NULL,
  `Telefono` varchar(20) DEFAULT NULL,
  `Pregunta` varchar(100) DEFAULT NULL,
  `Respuesta` varchar(50) DEFAULT NULL,
  `EmailAux` varchar(100) DEFAULT NULL,
  `Ban` tinyint(1) DEFAULT '0',
  `Nacimiento` date DEFAULT NULL,
  PRIMARY KEY (`Id`),
  UNIQUE KEY `Nombre` (`Nombre`)
) ENGINE=MyISAM AUTO_INCREMENT=1325 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `fortalezas`
--

DROP TABLE IF EXISTS `fortalezas`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `fortalezas` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `Nombre` varchar(50) DEFAULT NULL,
  `IdClan` int(11) DEFAULT NULL,
  `Fecha` datetime DEFAULT NULL,
  `X` int(11) DEFAULT NULL,
  `Y` int(11) DEFAULT NULL,
  `SpawnX` int(11) DEFAULT '0',
  `SpawnY` int(11) DEFAULT NULL,
  `NPCRey` int(11) DEFAULT NULL,
  PRIMARY KEY (`Id`),
  UNIQUE KEY `Id` (`Id`)
) ENGINE=MyISAM AUTO_INCREMENT=3 DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `penas`
--

DROP TABLE IF EXISTS `penas`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `penas` (
  `Id` tinyint(20) NOT NULL AUTO_INCREMENT,
  `IdPj` bigint(20) DEFAULT NULL,
  `Razon` varchar(100) DEFAULT NULL,
  `Fecha` datetime NOT NULL,
  `IdGM` bigint(20) DEFAULT NULL,
  `Tiempo` int(11) DEFAULT '0',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `Id` (`Id`)
) ENGINE=MyISAM AUTO_INCREMENT=128 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `pjs`
--

DROP TABLE IF EXISTS `pjs`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `pjs` (
  `Id` bigint(20) unsigned NOT NULL AUTO_INCREMENT,
  `IdAccount` bigint(20) DEFAULT '0',
  `Nombre` varchar(20) DEFAULT NULL,
  `Genero` tinyint(4) DEFAULT NULL,
  `Raza` tinyint(4) DEFAULT NULL,
  `Hogar` tinyint(4) DEFAULT NULL,
  `Clase` tinyint(4) DEFAULT NULL,
  `Heading` tinyint(4) DEFAULT NULL,
  `Head` int(11) DEFAULT NULL,
  `Body` int(11) DEFAULT NULL,
  `Arma` int(11) DEFAULT NULL,
  `Escudo` int(11) DEFAULT NULL,
  `Casco` int(11) DEFAULT NULL,
  `Uptime` bigint(20) DEFAULT NULL,
  `LastIP` varchar(15) DEFAULT NULL,
  `LastConnect` datetime DEFAULT NULL,
  `Map` tinyint(4) DEFAULT NULL,
  `X` int(11) DEFAULT NULL,
  `Y` int(11) DEFAULT NULL,
  `Logged` tinyint(1) DEFAULT NULL,
  `Muerto` tinyint(1) DEFAULT NULL,
  `Escondido` tinyint(1) DEFAULT NULL,
  `Hambre` tinyint(1) DEFAULT NULL,
  `Sed` tinyint(1) DEFAULT NULL,
  `Desnudo` tinyint(1) DEFAULT NULL,
  `Ban` tinyint(1) DEFAULT NULL,
  `Navegando` tinyint(1) DEFAULT NULL,
  `Envenenado` tinyint(1) DEFAULT NULL,
  `Paralizado` tinyint(1) DEFAULT NULL,
  `PerteneceReal` tinyint(1) DEFAULT NULL,
  `PerteneceCaos` tinyint(1) DEFAULT NULL,
  `Pena` int(11) DEFAULT NULL,
  `EjercitoReal` tinyint(1) DEFAULT NULL,
  `EjercitoCaos` tinyint(1) DEFAULT NULL,
  `CiudMatados` tinyint(1) DEFAULT NULL,
  `CrimMatados` tinyint(1) DEFAULT NULL,
  `rArCaos` tinyint(1) DEFAULT NULL,
  `rArReal` tinyint(1) DEFAULT NULL,
  `rExCaos` tinyint(1) DEFAULT NULL,
  `rExReal` tinyint(1) DEFAULT NULL,
  `recCaos` int(11) DEFAULT NULL,
  `recReal` int(11) DEFAULT NULL,
  `Reenlistadas` int(11) DEFAULT NULL,
  `NivelIngreso` int(11) DEFAULT NULL,
  `FechaIngreso` date DEFAULT NULL,
  `MatadosIngreso` int(11) DEFAULT NULL,
  `NextRecompensa` int(11) DEFAULT NULL,
  `At1` int(11) DEFAULT NULL,
  `At2` int(11) DEFAULT NULL,
  `At3` int(11) DEFAULT NULL,
  `At4` int(11) DEFAULT NULL,
  `At5` int(11) DEFAULT NULL,
  `Sk1` int(11) DEFAULT NULL,
  `Sk2` int(11) DEFAULT NULL,
  `Sk3` int(11) DEFAULT NULL,
  `Sk4` int(11) DEFAULT NULL,
  `Sk5` int(11) DEFAULT NULL,
  `Sk6` int(11) DEFAULT NULL,
  `Sk7` int(11) DEFAULT NULL,
  `Sk8` int(11) DEFAULT NULL,
  `Sk9` int(11) DEFAULT NULL,
  `Sk10` int(11) DEFAULT NULL,
  `Sk11` int(11) DEFAULT NULL,
  `Sk12` int(11) DEFAULT NULL,
  `Sk13` int(11) DEFAULT NULL,
  `Sk14` int(11) DEFAULT NULL,
  `Sk15` int(11) DEFAULT NULL,
  `Sk16` int(11) DEFAULT NULL,
  `Sk17` int(11) DEFAULT NULL,
  `Sk18` int(11) DEFAULT NULL,
  `Sk19` int(11) DEFAULT NULL,
  `Sk20` int(11) DEFAULT NULL,
  `Email` varchar(100) DEFAULT NULL,
  `Gld` bigint(20) DEFAULT NULL,
  `Banco` bigint(20) DEFAULT NULL,
  `MaxHP` int(11) DEFAULT NULL,
  `MinHP` int(11) DEFAULT NULL,
  `MaxSTA` int(11) DEFAULT NULL,
  `MinSTA` int(11) DEFAULT NULL,
  `MaxMAN` int(11) DEFAULT NULL,
  `MinMAN` int(11) DEFAULT NULL,
  `MaxHIT` int(11) DEFAULT NULL,
  `MinHIT` int(11) DEFAULT NULL,
  `MaxAGU` int(11) DEFAULT NULL,
  `MinAGU` int(11) DEFAULT NULL,
  `MaxHAM` int(11) DEFAULT NULL,
  `MinHAM` int(11) DEFAULT NULL,
  `SkillPtsLibres` int(11) DEFAULT NULL,
  `EXP` bigint(20) DEFAULT NULL,
  `ELV` int(11) DEFAULT NULL,
  `ELU` bigint(20) DEFAULT NULL,
  `UserMuertes` bigint(20) DEFAULT NULL,
  `NpcsMuertes` bigint(20) DEFAULT NULL,
  `WeaponEqpSlot` int(11) DEFAULT NULL,
  `ArmourEqpSlot` int(11) DEFAULT NULL,
  `CascoEqpSlot` int(11) DEFAULT NULL,
  `EscudoEqpSlot` int(11) DEFAULT NULL,
  `BarcoSlot` int(11) DEFAULT NULL,
  `MunicionSlot` int(11) DEFAULT NULL,
  `AnilloSlot` int(11) DEFAULT NULL,
  `Rep_Asesino` bigint(20) DEFAULT NULL,
  `Rep_Bandido` bigint(20) DEFAULT NULL,
  `Rep_Burguesia` bigint(20) DEFAULT NULL,
  `Rep_Ladrones` bigint(20) DEFAULT NULL,
  `Rep_Nobles` bigint(20) DEFAULT NULL,
  `Rep_Plebe` bigint(20) DEFAULT NULL,
  `Rep_Promedio` bigint(20) DEFAULT NULL,
  `NroMascotas` int(11) DEFAULT NULL,
  `Masc1` int(11) DEFAULT NULL,
  `Masc2` int(11) DEFAULT NULL,
  `Masc3` int(11) DEFAULT NULL,
  `TrainningTime` bigint(20) DEFAULT '0',
  `H1` int(11) DEFAULT NULL,
  `H2` int(11) DEFAULT NULL,
  `H3` int(11) DEFAULT NULL,
  `H4` int(11) DEFAULT NULL,
  `H5` int(11) DEFAULT NULL,
  `H6` int(11) DEFAULT NULL,
  `H7` int(11) DEFAULT NULL,
  `H8` int(11) DEFAULT NULL,
  `H9` int(11) DEFAULT NULL,
  `H10` int(11) DEFAULT NULL,
  `H11` int(11) DEFAULT NULL,
  `H12` int(11) DEFAULT NULL,
  `H13` int(11) DEFAULT NULL,
  `H14` int(11) DEFAULT NULL,
  `H15` int(11) DEFAULT NULL,
  `H16` int(11) DEFAULT NULL,
  `H17` int(11) DEFAULT NULL,
  `H18` int(11) DEFAULT NULL,
  `H19` int(11) DEFAULT NULL,
  `H20` int(11) DEFAULT NULL,
  `H21` int(11) DEFAULT NULL,
  `H22` int(11) DEFAULT NULL,
  `H23` int(11) DEFAULT NULL,
  `H24` int(11) DEFAULT NULL,
  `H25` int(11) DEFAULT NULL,
  `H26` int(11) DEFAULT NULL,
  `H27` int(11) DEFAULT NULL,
  `H28` int(11) DEFAULT NULL,
  `H29` int(11) DEFAULT NULL,
  `H30` int(11) DEFAULT NULL,
  `H31` int(11) DEFAULT NULL,
  `H32` int(11) DEFAULT NULL,
  `H33` int(11) DEFAULT NULL,
  `H34` int(11) DEFAULT NULL,
  `H35` int(11) DEFAULT NULL,
  `BanObj1` int(11) DEFAULT NULL,
  `BanCant1` int(11) DEFAULT NULL,
  `BanObj2` int(11) DEFAULT NULL,
  `BanCant2` int(11) DEFAULT NULL,
  `BanObj3` int(11) DEFAULT NULL,
  `BanCant3` int(11) DEFAULT NULL,
  `BanObj4` int(11) DEFAULT NULL,
  `BanCant4` int(11) DEFAULT NULL,
  `BanObj5` int(11) DEFAULT NULL,
  `BanCant5` int(11) DEFAULT NULL,
  `BanObj6` int(11) DEFAULT NULL,
  `BanCant6` int(11) DEFAULT NULL,
  `BanObj7` int(11) DEFAULT NULL,
  `BanCant7` int(11) DEFAULT NULL,
  `BanObj8` int(11) DEFAULT NULL,
  `BanCant8` int(11) DEFAULT NULL,
  `BanObj9` int(11) DEFAULT NULL,
  `BanCant9` int(11) DEFAULT NULL,
  `BanObj10` int(11) DEFAULT NULL,
  `BanCant10` int(11) DEFAULT NULL,
  `BanObj11` int(11) DEFAULT NULL,
  `BanCant11` int(11) DEFAULT NULL,
  `BanObj12` int(11) DEFAULT NULL,
  `BanCant12` int(11) DEFAULT NULL,
  `BanObj13` int(11) DEFAULT NULL,
  `BanCant13` int(11) DEFAULT NULL,
  `BanObj14` int(11) DEFAULT NULL,
  `BanCant14` int(11) DEFAULT NULL,
  `BanObj15` int(11) DEFAULT NULL,
  `BanCant15` int(11) DEFAULT NULL,
  `BanObj16` int(11) DEFAULT NULL,
  `BanCant16` int(11) DEFAULT NULL,
  `BanObj17` int(11) DEFAULT NULL,
  `BanCant17` int(11) DEFAULT NULL,
  `BanObj18` int(11) DEFAULT NULL,
  `BanCant18` int(11) DEFAULT NULL,
  `BanObj19` int(11) DEFAULT NULL,
  `BanCant19` int(11) DEFAULT NULL,
  `BanObj20` int(11) DEFAULT NULL,
  `BanCant20` int(11) DEFAULT NULL,
  `BanObj21` int(11) DEFAULT NULL,
  `BanCant21` int(11) DEFAULT NULL,
  `BanObj22` int(11) DEFAULT NULL,
  `BanCant22` int(11) DEFAULT NULL,
  `BanObj23` int(11) DEFAULT NULL,
  `BanCant23` int(11) DEFAULT NULL,
  `BanObj24` int(11) DEFAULT NULL,
  `BanCant24` int(11) DEFAULT NULL,
  `BanObj25` int(11) DEFAULT NULL,
  `BanCant25` int(11) DEFAULT NULL,
  `BanObj26` int(11) DEFAULT NULL,
  `BanCant26` int(11) DEFAULT NULL,
  `BanObj27` int(11) DEFAULT NULL,
  `BanCant27` int(11) DEFAULT NULL,
  `BanObj28` int(11) DEFAULT NULL,
  `BanCant28` int(11) DEFAULT NULL,
  `BanObj29` int(11) DEFAULT NULL,
  `BanCant29` int(11) DEFAULT NULL,
  `BanObj30` int(11) DEFAULT NULL,
  `BanCant30` int(11) DEFAULT NULL,
  `BanObj31` int(11) DEFAULT NULL,
  `BanCant31` int(11) DEFAULT NULL,
  `BanObj32` int(11) DEFAULT NULL,
  `BanCant32` int(11) DEFAULT NULL,
  `BanObj33` int(11) DEFAULT NULL,
  `BanCant33` int(11) DEFAULT NULL,
  `BanObj34` int(11) DEFAULT NULL,
  `BanCant34` int(11) DEFAULT NULL,
  `BanObj35` int(11) DEFAULT NULL,
  `BanCant35` int(11) DEFAULT NULL,
  `BanObj36` int(11) DEFAULT NULL,
  `BanCant36` int(11) DEFAULT NULL,
  `BanObj37` int(11) DEFAULT NULL,
  `BanCant37` int(11) DEFAULT NULL,
  `BanObj38` int(11) DEFAULT NULL,
  `BanCant38` int(11) DEFAULT NULL,
  `BanObj39` int(11) DEFAULT NULL,
  `BanCant39` int(11) DEFAULT NULL,
  `BanObj40` int(11) DEFAULT NULL,
  `BanCant40` int(11) DEFAULT NULL,
  `InvObj1` int(11) DEFAULT NULL,
  `InvCant1` int(11) DEFAULT NULL,
  `InvEqp1` tinyint(1) DEFAULT NULL,
  `InvObj2` int(11) DEFAULT NULL,
  `InvCant2` int(11) DEFAULT NULL,
  `InvEqp2` tinyint(1) DEFAULT NULL,
  `InvObj3` int(11) DEFAULT NULL,
  `InvCant3` int(11) DEFAULT NULL,
  `InvEqp3` tinyint(1) DEFAULT NULL,
  `InvObj4` int(11) DEFAULT NULL,
  `InvCant4` int(11) DEFAULT NULL,
  `InvEqp4` tinyint(1) DEFAULT NULL,
  `InvObj5` int(11) DEFAULT NULL,
  `InvCant5` int(11) DEFAULT NULL,
  `InvEqp5` tinyint(1) DEFAULT NULL,
  `InvObj6` int(11) DEFAULT NULL,
  `InvCant6` int(11) DEFAULT NULL,
  `InvEqp6` tinyint(1) DEFAULT NULL,
  `InvObj7` int(11) DEFAULT NULL,
  `InvCant7` int(11) DEFAULT NULL,
  `InvEqp7` tinyint(1) DEFAULT NULL,
  `InvObj8` int(11) DEFAULT NULL,
  `InvCant8` int(11) DEFAULT NULL,
  `InvEqp8` tinyint(1) DEFAULT NULL,
  `InvObj9` int(11) DEFAULT NULL,
  `InvCant9` int(11) DEFAULT NULL,
  `InvEqp9` tinyint(1) DEFAULT NULL,
  `InvObj10` int(11) DEFAULT NULL,
  `InvCant10` int(11) DEFAULT NULL,
  `InvEqp10` tinyint(1) DEFAULT NULL,
  `InvObj11` int(11) DEFAULT NULL,
  `InvCant11` int(11) DEFAULT NULL,
  `InvEqp11` tinyint(1) DEFAULT NULL,
  `InvObj12` int(11) DEFAULT NULL,
  `InvCant12` int(11) DEFAULT NULL,
  `InvEqp12` tinyint(1) DEFAULT NULL,
  `InvObj13` int(11) DEFAULT NULL,
  `InvCant13` int(11) DEFAULT NULL,
  `InvEqp13` tinyint(1) DEFAULT NULL,
  `InvObj14` int(11) DEFAULT NULL,
  `InvCant14` int(11) DEFAULT NULL,
  `InvEqp14` tinyint(1) DEFAULT NULL,
  `InvObj15` int(11) DEFAULT NULL,
  `InvCant15` int(11) DEFAULT NULL,
  `InvEqp15` tinyint(1) DEFAULT NULL,
  `InvObj16` int(11) DEFAULT NULL,
  `InvCant16` int(11) DEFAULT NULL,
  `InvEqp16` tinyint(1) DEFAULT NULL,
  `InvObj17` int(11) DEFAULT NULL,
  `InvCant17` int(11) DEFAULT NULL,
  `InvEqp17` tinyint(1) DEFAULT NULL,
  `InvObj18` int(11) DEFAULT NULL,
  `InvCant18` int(11) DEFAULT NULL,
  `InvEqp18` tinyint(1) DEFAULT NULL,
  `InvObj19` int(11) DEFAULT NULL,
  `InvCant19` int(11) DEFAULT NULL,
  `InvEqp19` tinyint(1) DEFAULT NULL,
  `InvObj20` int(11) DEFAULT NULL,
  `InvCant20` int(11) DEFAULT NULL,
  `InvEqp20` tinyint(1) DEFAULT NULL,
  `InvCantidadItems` int(11) DEFAULT NULL,
  `BanCantidadItems` int(11) DEFAULT NULL,
  `GuildIndex` int(11) DEFAULT '0',
  `Descripcion` varchar(100) DEFAULT NULL,
  `Creado` datetime DEFAULT NULL,
  `BannedBy` varchar(20) DEFAULT NULL,
  `Voto` int(11) DEFAULT '0',
  `AspiranteA` int(11) DEFAULT '0',
  `MotivoRechazo` varchar(100) DEFAULT '',
  `Pedidos` varchar(400) DEFAULT '0',
  `Miembro` varchar(400) DEFAULT NULL,
  `Extra` varchar(100) DEFAULT NULL,
  `Penas` int(11) DEFAULT '0',
  `BanTime` date DEFAULT '2000-01-01',
  PRIMARY KEY (`Id`),
  UNIQUE KEY `Nombre` (`Nombre`)
) ENGINE=MyISAM AUTO_INCREMENT=2355 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;
/*!40103 SET TIME_ZONE=@OLD_TIME_ZONE */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;

-- Dump completed on 2020-07-14 15:50:55

LOCK TABLES `fortalezas` WRITE;
/*!40000 ALTER TABLE `fortalezas` DISABLE KEYS */;
INSERT INTO `fortalezas` VALUES (1,'Oeste',11,'2018-12-15 14:50:15',40,1460,18,1304,620),(2,'Este',11,'2018-12-15 14:47:58',1055,1460,1080,1252,619);
/*!40000 ALTER TABLE `fortalezas` ENABLE KEYS */;
UNLOCK TABLES;