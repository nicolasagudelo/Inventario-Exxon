-- MySQL dump 10.13  Distrib 5.7.18, for Win32 (AMD64)
--
-- Host: localhost    Database: bd_productos
-- ------------------------------------------------------
-- Server version	5.7.18-log

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
-- Table structure for table `categorias`
--

DROP TABLE IF EXISTS `categorias`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `categorias` (
  `ID_Categoria` int(11) NOT NULL COMMENT 'ID de la Categoria',
  `Nombre` varchar(30) NOT NULL COMMENT 'Nombre de la Categoria',
  PRIMARY KEY (`ID_Categoria`),
  UNIQUE KEY `Nombre_UNIQUE` (`Nombre`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `categorias`
--

LOCK TABLES `categorias` WRITE;
/*!40000 ALTER TABLE `categorias` DISABLE KEYS */;
INSERT INTO `categorias` VALUES (18,'ASTM DIESEL'),(19,'ASTM GASOLINA'),(17,'ASTM JET A'),(6,'BASICAS B0E-B20E'),(8,'BASICAS B21E-B99E'),(9,'BASICAS GASOLINA'),(2,'BASICAS JET A'),(5,'COMPLETAS B0E-B20E'),(7,'COMPLETAS B21E-B99E'),(12,'COMPLETAS BIODIESEL (B100)'),(13,'COMPLETAS ETANOL'),(10,'COMPLETAS GASOLINA'),(11,'COMPLETAS GASOLINA OXIGENADA'),(1,'COMPLETAS JET A'),(15,'DENSIMETRO DIGITAL'),(14,'HIDROMETROS'),(16,'MICROBIOLOGIA'),(20,'TODOS');
/*!40000 ALTER TABLE `categorias` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `categoriaxproducto`
--

DROP TABLE IF EXISTS `categoriaxproducto`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `categoriaxproducto` (
  `ID_Producto` smallint(6) NOT NULL COMMENT 'ID del producto',
  `ID_Categoria` int(11) NOT NULL COMMENT 'ID de la Categoria',
  PRIMARY KEY (`ID_Producto`,`ID_Categoria`),
  KEY `FK_CategoriaXProductoCategoria` (`ID_Categoria`),
  CONSTRAINT `categoriaxproducto_ibfk_1` FOREIGN KEY (`ID_Producto`) REFERENCES `tipo_productos` (`Tipo_Producto_ID`),
  CONSTRAINT `categoriaxproducto_ibfk_2` FOREIGN KEY (`ID_Categoria`) REFERENCES `categorias` (`ID_Categoria`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `categoriaxproducto`
--

LOCK TABLES `categoriaxproducto` WRITE;
/*!40000 ALTER TABLE `categoriaxproducto` DISABLE KEYS */;
INSERT INTO `categoriaxproducto` VALUES (4,1),(4,2),(3,5),(5,5),(3,6),(5,6),(3,7),(5,7),(3,8),(5,8),(6,9),(7,9),(6,10),(7,10),(8,11),(9,11),(2,12),(1,13),(10,14),(11,15),(12,15),(13,15),(2,16),(3,16),(4,16),(5,16),(4,17),(5,18),(6,19),(7,19),(3,20),(14,20);
/*!40000 ALTER TABLE `categoriaxproducto` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `clientes`
--

DROP TABLE IF EXISTS `clientes`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `clientes` (
  `ClienteID` smallint(6) NOT NULL COMMENT 'Id del Cliente',
  `Nombre` varchar(50) NOT NULL COMMENT 'Nombre del Cliente',
  `Direccion` varchar(150) NOT NULL,
  `Contraseña` varchar(20) NOT NULL DEFAULT '123',
  PRIMARY KEY (`ClienteID`),
  UNIQUE KEY `Nombre_UNIQUE` (`Nombre`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `clientes`
--

LOCK TABLES `clientes` WRITE;
/*!40000 ALTER TABLE `clientes` DISABLE KEYS */;
INSERT INTO `clientes` VALUES (1,'PLANTA BOGOTA','Direccion Prueba','123'),(2,'PLANTA MANCILLA','Direccion Prueba2','123'),(3,'PLANTA MEDELLIN','Direccion Prueba3','123'),(4,'PLANTA LA DORADA','Direccion Prueba4','123'),(5,'PLANTA GALAPA','Direccion Prueba5','123'),(6,'PLANTA CARTAGENA','Direccion Prueba6','123'),(7,'PLANTA BUCARAMANGA','Direccion Prueba7','123'),(8,'PLANTA CARTAGO','Direccion Prueba8','123'),(9,'PLANTA GUALANDAY','Direccion Prueba9','123'),(10,'PLANTA YUMBO','Direccion Prueba10','123'),(11,'PLANTA CONJUNTA BUENAVENTURA','Direccion Prueba11','123'),(12,'PLANTA CHEVRON BOGOTA','Direccion Prueba12','123'),(13,'PLANTA CHEVRON CARTAGENA','Direccion Prueba13','123');
/*!40000 ALTER TABLE `clientes` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `datosproducto`
--

DROP TABLE IF EXISTS `datosproducto`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `datosproducto` (
  `DatoID` int(11) NOT NULL COMMENT 'ID del dato ingresado',
  `ProductoID` varchar(20) NOT NULL COMMENT 'ID del producto',
  `ID_Prueba` varchar(20) NOT NULL COMMENT 'ID de la prueba',
  `Dato_1` varchar(20) DEFAULT NULL,
  `Dato_2` varchar(20) DEFAULT NULL,
  `Dato_3` varchar(20) DEFAULT NULL,
  `Dato_4` varchar(20) DEFAULT NULL,
  `Dato_5` varchar(20) DEFAULT NULL,
  `Dato_6` varchar(20) DEFAULT NULL,
  `Dato_7` varchar(20) DEFAULT NULL,
  `Dato_8` varchar(20) DEFAULT NULL,
  `Dato_9` varchar(20) DEFAULT NULL,
  `Dato_10` varchar(20) DEFAULT NULL,
  `Dato_11` varchar(20) DEFAULT NULL,
  `Dato_12` varchar(20) DEFAULT NULL,
  `Dato_13` varchar(20) DEFAULT NULL,
  `Dato_14` varchar(20) DEFAULT NULL,
  `Dato_15` varchar(20) DEFAULT NULL,
  `Dato_16` varchar(20) DEFAULT NULL,
  `Dato_17` varchar(20) DEFAULT NULL,
  `Dato_18` varchar(20) DEFAULT NULL,
  `Dato_19` varchar(20) DEFAULT NULL,
  `Dato_20` varchar(20) DEFAULT NULL,
  PRIMARY KEY (`DatoID`),
  KEY `DatosProducto_ibfk_1` (`ID_Prueba`),
  KEY `DatosProducto_ibfk_2` (`ProductoID`),
  CONSTRAINT `DatosProducto_ibfk_1` FOREIGN KEY (`ID_Prueba`) REFERENCES `pruebas` (`ID_Prueba`),
  CONSTRAINT `DatosProducto_ibfk_2` FOREIGN KEY (`ProductoID`) REFERENCES `productos` (`ProductoID`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `datosproducto`
--

LOCK TABLES `datosproducto` WRITE;
/*!40000 ALTER TABLE `datosproducto` DISABLE KEYS */;
INSERT INTO `datosproducto` VALUES (1,'2017-9','D-3242','0,0299','10,45','0,05','0,0141','100','6,85','0,15','30',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(2,'2017-12','D-1319-01','582','592','595','603','110','110','110','110','18,9','18,6','18,5','18,2','18,6',NULL,NULL,NULL,NULL,NULL,NULL,NULL),(3,'2017-12','D-1319-02','582','592','595','603','9','9','9','9','1,5','1,5','1,5','1,5','1,5',NULL,NULL,NULL,NULL,NULL,NULL,NULL),(4,'2017-12','D-1319-03','582','592','595','603','474','474','474','474','81,4','80,1','79,7','78,6','80,0',NULL,NULL,NULL,NULL,NULL,NULL,NULL),(5,'2017-12','D-1322','No','No','1','32','20','20','20','20','0,98','19,5',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(6,'2017-12','D-1840','0,4225','0,099','0,1','33,7','0,8014','1','1,266372928094','0,8014','1,01',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(7,'2017-12','D-3242','0,0299','10,45','0,05','0,0141','100','6,85','0,15','30','0,053',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(8,'2017-12','D-1322','No','No','1','32','20','20','20','20','0,98','19,5',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(9,'2017-13','D-56','30','75,1','37,5','44',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(10,'2017-13','D-93','35','75,32','39,5','46',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(11,'2017-14','D-4294-01','SI','SI','13,2','14,6','13,9',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(12,'2017-14','D-4294-02','NO','NO','10,5','8,6','9,55',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(13,'2017-14','D-4294-01','SI','SI','13,2',NULL,'13,2',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(14,'2017-14','D-4294-02','NO','NO','10,5',NULL,'10,5',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(15,'2017-15','D-3241-01','12345678','OBSERVACIONES','35','11:00','10:35','85','600','0,78','Caida Presion','11:30',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(16,'2017-15','D-3241-02','12345678','OBSERVACIONES','35','11:00','10:35','85','600','0,78','Rojo','11:30',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(17,'2017-16','D-86-01','32','1','Observaciones','NO','NO','NO','100',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(18,'2017-16','D-86-03','32','1','Observaciones','NO','NO','NO','110',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(19,'2017-16','D-86-07','32','1','Observaciones','NO','NO','NO','150',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(20,'2017-16','D-86-11','32','1','Observaciones','NO','NO','NO','190',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(21,'2017-16','D-86-12','32','1','Observaciones','NO','NO','NO','195',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(22,'2017-16','D-86-14','32','1','Observaciones','NO','NO','NO','110',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(23,'2017-16','D-86-17','32','1','Observaciones','NO','NO','NO','140',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(24,'2017-16','D-86-18','32','1','Observaciones','NO','NO','NO','150',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(25,'2017-16','D-86-22','32','1','Observaciones','NO','NO','NO','190',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(26,'2017-16','D-86-24','32','1','Observaciones','NO','NO','NO','200',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(27,'2017-16','D-86-25','32','1','Observaciones','NO','NO','NO','20',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(28,'2017-16','D-86-26','32','1','Observaciones','NO','NO','NO','10',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(29,'2017-16','D-86-27','32','1','Observaciones','NO','NO','NO','80',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(30,'2017-16','D-86-28','32','1','Observaciones','NO','NO','NO','20',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(32,'2017-18','D-4176','89,6','NO PASA','NO PASA','NO PASA','NO PASA','Observaciones',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(33,'2017-17','D-2386','0','-5','-50','-45,6','-45,5',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(34,'2017-19','D-1094-01','34','500','Observaciones','40',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(35,'2017-19','D-1094-02','34','500','Observaciones','2',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(36,'2017-19','D-1094-03','34','500','Observaciones','1',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(37,'2017-19','D-3948','NO','NO','34','32','-2','30','35',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(38,'2017-20','D-130-01','NO','NO','50',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(39,'2017-20','D-130-02','NO','NO','25',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(40,'2017-21','D-4953','SI','NO','SI','NO','15','20','25','45','310,25','30','40','42','82','565,50',NULL,NULL,NULL,NULL,NULL,NULL),(41,'2017-22','D-4815','1500','1000','40','60',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(42,'2017-23','D-1298','104','15','101,8','102,9','20','1,5','25',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(43,'2017-25','D-445','NO','SI','10',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(44,'2017-26','MMIN','565,5','40','610,7',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(45,'2017-28','EN14078','40','20','55','OBSERVACION','65',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(46,'2017-29','D-7321','1000','700','300','0','0','0','0','0','0','200','0','0','300','200','1500000',NULL,NULL,NULL,NULL,NULL),(47,'2017-29','EN12937','1250','Observaciones','82,974',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(48,'2017-30','D-525','11:30','Si','101','10','11,01',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(49,'2017-31','D-381','58,4645','58,4643','82,2685','82,2683','NO','NO','0,799999999998136','<1',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(50,'2017-32','D-381','58,4645','58,4643','82,2685','82,2683','58,4641','82,2684','NO','NO','0,799999999998136','1','-0,199999999978218','<0.5',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(51,'2017-33','D-3338','18,6','20','110','150','190','150','345,5330774','13,9','298,917',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(52,'R-2017-3a','D-1094-01','34','500','Observaciones','40',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(53,'R-2017-3a','D-1094-02','34','500','Observaciones','2',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(54,'R-2017-3a','D-1094-03','34','500','Observaciones','1',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(55,'R-2017-3a','D-1298','104','15','101,8','102,9','20','1,5','25',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(56,'R-2017-3a','D-130-01','NO','NO','50',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(57,'R-2017-3a','D-1319-01','582','592','595','603','110','110','110','110','18,9','18,6','18,5','18,2','18,6',NULL,NULL,NULL,NULL,NULL,NULL,NULL),(58,'R-2017-3a','D-1322','No','No','1','32','20','20','20','20','0,98','19,5',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(59,'R-2017-3a','D-1840','0,4225','0,099','0,1','33,7','0,8014','1','1,266372928094','0,8014','1,01',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(60,'R-2017-3a','D-2386','0','-5','-50','-45','-45,0',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(61,'R-2017-3a','D-3241-01','12345678','OBSERVACIONES','35','11:00','10:35','85','600','0,78','Caida Presion','11:30',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(62,'R-2017-3a','D-3241-02','12345678','OBSERVACIONES','35','11:00','10:35','85','600','0,78','Rojo','11:30',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(63,'R-2017-3a','D-3242','0,0299','10,45','0,05','0,0141','100','6,85','0,15','30','0,053',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(64,'R-2017-3a','D-3338','18,6','20','110','150','190','150','345,5330774','13,9','298,917',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(65,'R-2017-3a','D-381','58,4645','58,4643','82,2685','82,2683','NO','NO','0,799999999998136','<1',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(66,'R-2017-3a','D-3948','NO','NO','34','32','-2','30','35',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(67,'R-2017-3a','D-4176','89,6','NO PASA','NO PASA','NO PASA','NO PASA','Observaciones',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(68,'R-2017-3a','D-4294-01','SI','SI','13,2','14,6','13,9',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(69,'R-2017-3a','D-445','NO','SI','10',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(70,'R-2017-3a','D-56','30','75,1','37,5','44',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(71,'R-2017-3a','D-86-01','32','1','Observaciones','NO','NO','NO','100',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(72,'R-2017-3a','D-86-03','32','1','Observaciones','NO','NO','NO','110',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(73,'R-2017-3a','D-86-07','32','1','Observaciones','NO','NO','NO','150',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(74,'R-2017-3a','D-86-11','32','1','Observaciones','NO','NO','NO','190',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(75,'R-2017-3a','D-86-24','32','1','Observaciones','NO','NO','NO','200',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(76,'R-2017-3a','D-86-25','32','1','Observaciones','NO','NO','NO','20',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL),(77,'R-2017-3a','D-86-26','32','1','Observaciones','NO','NO','NO','10',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL);
/*!40000 ALTER TABLE `datosproducto` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `productos`
--

DROP TABLE IF EXISTS `productos`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `productos` (
  `ProductoID` varchar(20) NOT NULL COMMENT 'ID del producto',
  `ClienteID` smallint(6) NOT NULL COMMENT 'Id del Cliente proveniente de la tabla Clientes',
  `Tipo_Producto_ID` smallint(6) NOT NULL COMMENT 'ID del tipo de productos proveniente de la tabla de tipo de productos',
  `Observaciones_Cliente` varchar(500) DEFAULT NULL COMMENT 'Observaciones del producto si las hay',
  `Fecha_Registro` date NOT NULL,
  `Fecha_Entrada` date DEFAULT NULL COMMENT 'Fecha de entrada al laboratorio',
  `Hora_Entrada` time DEFAULT NULL,
  `Fecha_Limite` date DEFAULT NULL COMMENT 'Fecha_limite para generar el reporte',
  `Estado` varchar(20) NOT NULL COMMENT 'Estado del producto',
  `Observaciones_Analista` varchar(500) DEFAULT NULL COMMENT 'Observaciones del producto si las hay',
  `Fecha_Reporte` date DEFAULT NULL,
  `Observaciones_Administrador` varchar(500) DEFAULT NULL COMMENT 'Observaciones del producto si las hay',
  `UsuarioID` smallint(6) DEFAULT NULL,
  `ID_Muestra` varchar(50) DEFAULT NULL,
  `Tanque` varchar(50) DEFAULT NULL,
  `Lote` varchar(50) DEFAULT NULL,
  `ATN` varchar(50) DEFAULT NULL,
  `TipodePrueba` tinyint(4) DEFAULT NULL,
  `PortBioD` int(11) DEFAULT NULL,
  PRIMARY KEY (`ProductoID`),
  KEY `ClienteID` (`ClienteID`),
  KEY `Tipo_Producto_ID` (`Tipo_Producto_ID`),
  KEY `productos_ibfk_3_idx` (`UsuarioID`),
  CONSTRAINT `productos_ibfk_1` FOREIGN KEY (`ClienteID`) REFERENCES `clientes` (`ClienteID`),
  CONSTRAINT `productos_ibfk_2` FOREIGN KEY (`Tipo_Producto_ID`) REFERENCES `tipo_productos` (`Tipo_Producto_ID`),
  CONSTRAINT `productos_ibfk_3` FOREIGN KEY (`UsuarioID`) REFERENCES `usuarios` (`UsuarioID`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `productos`
--

LOCK TABLES `productos` WRITE;
/*!40000 ALTER TABLE `productos` DISABLE KEYS */;
INSERT INTO `productos` VALUES ('2017-1',1,2,'','2017-11-01','2017-11-03','15:25:20','2017-11-23','Pendiente',NULL,NULL,NULL,NULL,'','12','',NULL,NULL,NULL),('2017-10',1,4,NULL,'2017-11-08','2017-11-08','10:32:37','2017-11-23','Pendiente','',NULL,NULL,NULL,'','12','',NULL,NULL,NULL),('2017-11',1,4,NULL,'2017-11-10','2017-11-10','08:26:50','2017-11-13','Revisado','','2017-11-10',NULL,5,'','12','',NULL,NULL,NULL),('2017-12',1,14,NULL,'2017-11-10','2017-11-10','09:13:54','2017-11-22','Revisado','','2017-11-10','Repetir D-1322',2,'','123','',NULL,NULL,NULL),('2017-13',1,14,NULL,'2017-11-10','2017-11-10','13:19:31','2017-11-22','Revisado','','2017-11-10',NULL,5,'','tanque 1232','',NULL,NULL,NULL),('2017-14',1,14,NULL,'2017-11-10','2017-11-10','15:12:08','2017-11-22','Revisado','','2017-11-10','',2,'','prueba','',NULL,NULL,NULL),('2017-15',8,14,NULL,'2017-11-14','2017-11-14','08:47:45','2017-11-26','Revisado','','2017-11-14',NULL,5,'','prueba','',NULL,NULL,NULL),('2017-16',2,14,NULL,'2017-11-14','2017-11-14','09:42:29','2017-11-26','Revisado','','2017-11-14',NULL,5,'','Tanque Destilacion','',NULL,NULL,NULL),('2017-17',1,14,NULL,'2017-11-14','2017-11-14','10:27:24','2017-11-26','Revisado','','2017-11-14',NULL,5,'','1','',NULL,NULL,NULL),('2017-18',1,14,NULL,'2017-11-14','2017-11-14','10:53:49','2017-11-26','Revisado','','2017-11-14',NULL,2,'','123','',NULL,NULL,NULL),('2017-19',1,14,NULL,'2017-11-14','2017-11-14','12:02:26','2017-11-26','Revisado','','2017-11-14',NULL,5,'','prueba','',NULL,NULL,NULL),('2017-2',1,6,NULL,'2017-11-02','2017-11-02','10:12:50','2017-11-23','Pendiente','as',NULL,NULL,NULL,'as','as','as',NULL,NULL,NULL),('2017-20',1,14,NULL,'2017-11-14','2017-11-14','13:18:38','2017-11-26','Revisado','','2017-11-14',NULL,5,'','123','',NULL,NULL,NULL),('2017-21',1,14,NULL,'2017-11-14','2017-11-14','14:03:47','2017-11-26','Revisado','','2017-11-14',NULL,5,'','prueba','',NULL,NULL,NULL),('2017-22',1,14,NULL,'2017-11-14','2017-11-14','14:34:47','2017-11-26','Revisado','','2017-11-14',NULL,5,'','test','',NULL,NULL,NULL),('2017-23',1,14,NULL,'2017-11-14','2017-11-14','15:03:36','2017-11-26','Revisado','','2017-11-14',NULL,5,'','789','',NULL,NULL,NULL),('2017-24',1,14,NULL,'2017-11-14','2017-11-14','15:27:40','2017-11-26','Revisado','','2017-11-14',NULL,5,'','prueba','',NULL,NULL,NULL),('2017-25',1,14,NULL,'2017-11-14','2017-11-14','15:44:50','2017-11-26','Revisado','','2017-11-14',NULL,5,'','1231','',NULL,NULL,NULL),('2017-26',1,14,NULL,'2017-11-15','2017-11-15','08:19:11','2017-11-27','Revisado','','2017-11-15',NULL,5,'','654','',NULL,NULL,NULL),('2017-27',1,14,NULL,'2017-11-15','2017-11-15','08:19:11','2017-11-27','Revisado','','2017-11-15',NULL,5,'','654','',NULL,NULL,NULL),('2017-28',1,14,NULL,'2017-11-15','2017-11-15','09:04:32','2017-11-27','Revisado','','2017-11-15',NULL,5,'','789','',NULL,NULL,NULL),('2017-29',1,14,NULL,'2017-11-15','2017-11-15','10:24:37','2017-11-27','Revisado','','2017-11-15',NULL,5,'','asdas','',NULL,NULL,NULL),('2017-3',1,4,NULL,'2017-11-02','2017-11-02','10:13:19','2017-11-10','Aprobado','Observaciones Analista','2017-11-02','Admon',5,'Jet','Tanque 2','Lote 3','Prueba',2,NULL),('2017-30',1,14,NULL,'2017-11-15','2017-11-15','10:56:42','2017-11-27','Revisado','','2017-11-15',NULL,5,'','123123','',NULL,NULL,NULL),('2017-31',1,4,NULL,'2017-11-15','2017-11-15','13:20:54','2017-11-18','Revisado','','2017-11-15',NULL,5,'','456','',NULL,NULL,NULL),('2017-32',1,6,NULL,'2017-11-15','2017-11-15','13:21:14','2017-11-23','Revisado','','2017-11-15',NULL,5,'','Origen','',NULL,NULL,NULL),('2017-33',1,14,NULL,'2017-11-15','2017-11-15','14:26:45','2017-11-27','Revisado','','2017-11-15',NULL,5,'','awdsas','',NULL,NULL,NULL),('2017-34',1,5,NULL,'2017-11-15','2017-11-15','15:57:38','2017-11-25','Pendiente','',NULL,NULL,NULL,'','12312','',NULL,NULL,NULL),('2017-35',1,14,NULL,'2017-11-15','2017-11-15','15:58:21','2017-11-27','Pendiente','',NULL,NULL,NULL,'','LastTest','',NULL,NULL,NULL),('2017-4',6,3,NULL,'2017-11-02','2017-11-02','10:14:06','2017-11-23','Pendiente','Observaciones Analista','2017-11-02','repetir d-1250',2,'BioD','CTE264','Lote 1','Prueba',2,50),('2017-5',6,3,NULL,'2017-11-02','2017-11-02','10:14:06','2017-11-23','Pendiente','Observaciones Analista','2017-11-02','repetir d-1250',2,'BioD','CTE264','Lote 1','Prueba',2,50),('2017-6',1,6,NULL,'2017-11-03','2017-11-03','10:20:14','2017-11-11','Aprobado','Observaciones','2017-11-03','Observaciones Admin',2,'Corriente','123','123','Planta Bogota',2,NULL),('2017-7',1,4,'','2017-11-03','2017-11-03','15:48:59','2017-11-06','Aprobado','','2017-11-03','T ambiente',5,'','tanque 1','','Planta Bogota',1,NULL),('2017-8',1,4,NULL,'2017-11-08','2017-11-08','09:15:16','2017-11-23','Pendiente','observaciones',NULL,NULL,NULL,'idmuestra','12','13',NULL,NULL,NULL),('2017-9',1,4,NULL,'2017-11-08','2017-11-08','10:15:19','2017-11-11','Aprobado','','2017-11-08','obs',5,'1123','Tanque 123','lote 123','Planta Bogota',3,NULL),('R-2017-3a',1,4,NULL,'2017-11-15','2017-11-15','15:56:34','2017-11-18','Revisado','','2017-11-15',NULL,5,'Jet','Tanque 2','Lote 3',NULL,NULL,NULL);
/*!40000 ALTER TABLE `productos` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `pruebas`
--

DROP TABLE IF EXISTS `pruebas`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `pruebas` (
  `ID_Prueba` varchar(20) NOT NULL COMMENT 'ID de la prueba',
  `Nombre` varchar(150) NOT NULL COMMENT 'Nombre de la prueba',
  PRIMARY KEY (`ID_Prueba`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `pruebas`
--

LOCK TABLES `pruebas` WRITE;
/*!40000 ALTER TABLE `pruebas` DISABLE KEYS */;
INSERT INTO `pruebas` VALUES ('D-1094-01','Cambio de volumen'),('D-1094-02','Clasificación Condiciones Interfase'),('D-1094-03','Clasificacion Grado de Separacion Interfase'),('D-1250','Densidad 15ºC, Kg/m3'),('D-1298','Gravedad API 60ºF'),('D-130-01','Corrosión L. Cobre,/100 °C/2h'),('D-130-02','Corrosión L. Cobre,/50 °C/3h'),('D-1319-01','Aromaticos'),('D-1319-02','Olefinas'),('D-1319-03','Saturados'),('D-1322','Punto de Humo, mm'),('D-1840','Naftalenos, %V'),('D-2386','Punto de Congelación, ºC'),('D-2500','Punto de Nube, °C'),('D-323','Presion Vapor REID  100 °F, kPa  o PSI'),('D-3241-01','Caida Presion,'),('D-3241-02','Codigo Color Tubo 260 °C'),('D-3242','Acidez Total, mg KOH/g'),('D-3338','Calor Neto Combustión, MJ/Kg'),('D-381','Goma Existente, mg/100mL'),('D-3948','MSP-A'),('D-4176','Apariencia'),('D-4294-01','Azufre, %m'),('D-4294-02','Azufre, mg/kg'),('D-445','Viscosidad   - 20 ºC, cSt'),('D-4815','% Etanol, Gasolina Oxigenada'),('D-4952','Prueba Doctor'),('D-4953','Presion Vapor REID  de Gasolina y mezclas de Gasolinas Oxigenadas100 °F, kPa  o PSI'),('D-525','Estabilidad Oxidacion / minutos'),('D-56','Punto de Chispa TAG, ºC'),('D-7321','Determinación de Contaminación Total'),('D-7978-01','Determinación del contenido de microorganismos aeróbicos CFU/L'),('D-7978-02','Determinación del contenido de microorganismos aeróbicos CFU/mL'),('D-86-01','Punto Inicial Ebullición, ºC'),('D-86-02','Temp. 5% Recobrado . ºC'),('D-86-03','Temp. 10% Recobrado . ºC'),('D-86-04','Temp. 20% Recobrado . ºC'),('D-86-05','Temp. 30% Recobrado . ºC'),('D-86-06','Temp. 40% Recobrado . ºC'),('D-86-07','Temp. 50% Recobrado ºC'),('D-86-08','Temp. 60% Recobrado . ºC'),('D-86-09','Temp. 70% Recobrado . ºC'),('D-86-10','Temp. 80% Recobrado . ºC'),('D-86-11','Temp. 90% Recobrado °C'),('D-86-12','Temp. 95%  Recobrado ºC'),('D-86-13','Temp. 5%  Volumen Evaporado. ºC'),('D-86-14','Temp. 10%  Volumen Evaporado. ºC'),('D-86-15','Temp. 20%  Volumen Evaporado. ºC'),('D-86-16','Temp. 30%  Volumen Evaporado. ºC'),('D-86-17','Temp. 40%  Volumen Evaporado. ºC'),('D-86-18','Temp. 50%  Volumen Evaporado. ºC'),('D-86-19','Temp. 60%  Volumen Evaporado. ºC'),('D-86-20','Temp. 70%  Volumen Evaporado. ºC'),('D-86-21','Temp. 80%  Volumen Evaporado. ºC'),('D-86-22','Temp. 90%  Volumen Evaporado. ºC'),('D-86-23','Temp. 95%  Volumen Evaporado. ºC'),('D-86-24','Punto Final Ebullición, ºC'),('D-86-25','Residuo Destilación.%'),('D-86-26','Pérdidas Destilación.%'),('D-86-27','Volumen Recogido %'),('D-86-28','Residuo  Volumen %'),('D-93','Punto de Chispa PENSKY MARTENS, ºC'),('D-97','Punto de Fluidez, °C'),('DM-01','Verificacion'),('DM-02','Calibracion'),('E-1064','Humedad por Titulación Coulometrica Karl Fisher % masa'),('EN12937','Determinacion de Humedad Karl Fisher %M.'),('EN14078','Determinacion de BX en Diesel/ACEM'),('EN14103-01','Contenido FAME\'s totales (%masa)'),('EN14103-02','Contenido de Ester del Acido Linolenico (%masa)'),('HM-01','Verificacion 2 Puntos'),('HM-02','Verificacion 3 Puntos'),('MMIN','Indice de Cierre de Vapor (ICV) kPa'),('VISUAL','Color');
/*!40000 ALTER TABLE `pruebas` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `pruebasxcategoria`
--

DROP TABLE IF EXISTS `pruebasxcategoria`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `pruebasxcategoria` (
  `ID_Prueba` varchar(20) NOT NULL COMMENT 'ID de la prueba',
  `ID_Categoria` int(11) NOT NULL COMMENT 'ID de la Categoria',
  PRIMARY KEY (`ID_Prueba`,`ID_Categoria`),
  KEY `FK_PruebasXCategoriaCategoria` (`ID_Categoria`),
  CONSTRAINT `pruebasxcategoria_ibfk_1` FOREIGN KEY (`ID_Prueba`) REFERENCES `pruebas` (`ID_Prueba`),
  CONSTRAINT `pruebasxcategoria_ibfk_2` FOREIGN KEY (`ID_Categoria`) REFERENCES `categorias` (`ID_Categoria`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `pruebasxcategoria`
--

LOCK TABLES `pruebasxcategoria` WRITE;
/*!40000 ALTER TABLE `pruebasxcategoria` DISABLE KEYS */;
INSERT INTO `pruebasxcategoria` VALUES ('D-1094-01',1),('D-1094-02',1),('D-1094-03',1),('D-1250',1),('D-1298',1),('D-130-01',1),('D-1319-01',1),('D-1322',1),('D-1840',1),('D-2386',1),('D-3241-01',1),('D-3241-02',1),('D-3242',1),('D-3338',1),('D-381',1),('D-3948',1),('D-4176',1),('D-4294-01',1),('D-445',1),('D-4952',1),('D-56',1),('D-86-01',1),('D-86-03',1),('D-86-07',1),('D-86-11',1),('D-86-24',1),('D-86-25',1),('D-86-26',1),('D-1094-01',2),('D-1094-02',2),('D-1094-03',2),('D-1250',2),('D-1298',2),('D-130-01',2),('D-2386',2),('D-381',2),('D-3948',2),('D-4176',2),('D-56',2),('D-86-01',2),('D-86-03',2),('D-86-07',2),('D-86-11',2),('D-86-24',2),('D-86-25',2),('D-86-26',2),('D-1250',5),('D-1298',5),('D-2500',5),('D-4176',5),('D-4294-01',5),('D-4294-02',5),('D-7321',5),('D-86-01',5),('D-86-03',5),('D-86-07',5),('D-86-11',5),('D-86-12',5),('D-86-24',5),('D-86-25',5),('D-86-26',5),('D-93',5),('D-97',5),('EN12937',5),('EN14078',5),('D-1250',6),('D-1298',6),('D-4176',6),('D-86-01',6),('D-86-03',6),('D-86-07',6),('D-86-11',6),('D-86-12',6),('D-86-24',6),('D-86-25',6),('D-86-26',6),('D-93',6),('D-2500',7),('D-4176',7),('D-4294-01',7),('D-4294-02',7),('D-7321',7),('D-97',7),('EN12937',7),('EN14078',7),('D-4176',8),('EN12937',8),('EN14078',8),('D-1250',9),('D-1298',9),('D-4176',9),('D-86-01',9),('D-86-14',9),('D-86-18',9),('D-86-22',9),('D-86-24',9),('D-86-28',9),('D-130-02',10),('D-1319-01',10),('D-1319-02',10),('D-1319-03',10),('D-381',10),('D-4953',10),('D-525',10),('D-86-01',10),('D-86-14',10),('D-86-18',10),('D-86-22',10),('D-86-28',10),('MMIN',10),('D-130-02',11),('D-381',11),('D-4815',11),('D-4953',11),('D-525',11),('D-86-01',11),('D-86-14',11),('D-86-18',11),('D-86-22',11),('D-86-28',11),('MMIN',11),('D-1250',12),('D-1298',12),('D-130-01',12),('D-2500',12),('D-4176',12),('D-7321',12),('D-97',12),('EN12937',12),('D-1250',13),('D-1298',13),('E-1064',13),('HM-01',14),('HM-02',14),('DM-01',15),('DM-02',15),('D-7978-01',16),('D-7978-02',16),('D-1250',17),('D-1298',17),('D-130-01',17),('D-1319-01',17),('D-1319-02',17),('D-1319-03',17),('D-1322',17),('D-1840',17),('D-2386',17),('D-3241-01',17),('D-3241-02',17),('D-3242',17),('D-3338',17),('D-381',17),('D-3948',17),('D-4294-01',17),('D-445',17),('D-56',17),('D-86-01',17),('D-86-02',17),('D-86-03',17),('D-86-04',17),('D-86-05',17),('D-86-06',17),('D-86-07',17),('D-86-08',17),('D-86-09',17),('D-86-10',17),('D-86-11',17),('D-86-12',17),('D-86-24',17),('D-86-25',17),('D-86-26',17),('D-1250',18),('D-1298',18),('D-130-02',18),('D-1319-01',18),('D-1319-02',18),('D-1319-03',18),('D-2500',18),('D-4294-01',18),('D-4294-02',18),('D-86-01',18),('D-86-02',18),('D-86-03',18),('D-86-04',18),('D-86-05',18),('D-86-06',18),('D-86-07',18),('D-86-08',18),('D-86-09',18),('D-86-10',18),('D-86-11',18),('D-86-12',18),('D-86-24',18),('D-86-25',18),('D-86-26',18),('D-93',18),('D-97',18),('D-1250',19),('D-1298',19),('D-130-02',19),('D-1319-01',19),('D-1319-02',19),('D-1319-03',19),('D-381',19),('D-4815',19),('D-525',19),('D-86-01',19),('D-86-13',19),('D-86-14',19),('D-86-15',19),('D-86-16',19),('D-86-17',19),('D-86-18',19),('D-86-19',19),('D-86-20',19),('D-86-21',19),('D-86-22',19),('D-86-23',19),('D-86-24',19),('D-86-25',19),('D-86-26',19),('D-86-27',19);
/*!40000 ALTER TABLE `pruebasxcategoria` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `pruebasxproducto`
--

DROP TABLE IF EXISTS `pruebasxproducto`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `pruebasxproducto` (
  `ProductoID` varchar(20) NOT NULL COMMENT 'Id del producto o muestra',
  `ID_Prueba` varchar(20) NOT NULL COMMENT 'Prueba a realizar sobre el producto o muestra',
  PRIMARY KEY (`ProductoID`,`ID_Prueba`),
  KEY `ProductoID` (`ProductoID`),
  KEY `ID_Prueba` (`ID_Prueba`),
  CONSTRAINT `FK_ProdRel` FOREIGN KEY (`ProductoID`) REFERENCES `productos` (`ProductoID`),
  CONSTRAINT `FK_PrueRel` FOREIGN KEY (`ID_Prueba`) REFERENCES `pruebas` (`ID_Prueba`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `pruebasxproducto`
--

LOCK TABLES `pruebasxproducto` WRITE;
/*!40000 ALTER TABLE `pruebasxproducto` DISABLE KEYS */;
INSERT INTO `pruebasxproducto` VALUES ('2017-10','D-3242'),('2017-11','D-1319-01'),('2017-11','D-1322'),('2017-11','D-1840'),('2017-11','D-3242'),('2017-12','D-1319-01'),('2017-12','D-1319-02'),('2017-12','D-1319-03'),('2017-12','D-1322'),('2017-12','D-1840'),('2017-12','D-3242'),('2017-13','D-56'),('2017-13','D-93'),('2017-14','D-4294-01'),('2017-14','D-4294-02'),('2017-15','D-3241-01'),('2017-15','D-3241-02'),('2017-16','D-86-01'),('2017-16','D-86-03'),('2017-16','D-86-07'),('2017-16','D-86-11'),('2017-16','D-86-12'),('2017-16','D-86-14'),('2017-16','D-86-17'),('2017-16','D-86-18'),('2017-16','D-86-22'),('2017-16','D-86-24'),('2017-16','D-86-25'),('2017-16','D-86-26'),('2017-16','D-86-27'),('2017-16','D-86-28'),('2017-17','D-2386'),('2017-18','D-4176'),('2017-19','D-1094-01'),('2017-19','D-1094-02'),('2017-19','D-1094-03'),('2017-19','D-3948'),('2017-2','D-1094-01'),('2017-2','D-1094-02'),('2017-2','D-1094-03'),('2017-2','D-1250'),('2017-2','D-1298'),('2017-2','D-130-01'),('2017-2','D-2386'),('2017-2','D-381'),('2017-2','D-3948'),('2017-2','D-4176'),('2017-2','D-56'),('2017-2','D-86-01'),('2017-2','D-86-03'),('2017-2','D-86-07'),('2017-2','D-86-11'),('2017-2','D-86-24'),('2017-2','D-86-25'),('2017-2','D-86-26'),('2017-20','D-130-01'),('2017-20','D-130-02'),('2017-21','D-4953'),('2017-22','D-4815'),('2017-23','D-1298'),('2017-24','D-1250'),('2017-24','D-4952'),('2017-24','E-1064'),('2017-25','D-445'),('2017-26','MMIN'),('2017-28','EN14078'),('2017-29','D-7321'),('2017-29','EN12937'),('2017-3','D-1094-01'),('2017-3','D-1094-02'),('2017-3','D-1094-03'),('2017-3','D-1250'),('2017-3','D-1298'),('2017-3','D-130-01'),('2017-3','D-1319-01'),('2017-3','D-1322'),('2017-3','D-1840'),('2017-3','D-2386'),('2017-3','D-3241-01'),('2017-3','D-3241-02'),('2017-3','D-3242'),('2017-3','D-3338'),('2017-3','D-381'),('2017-3','D-3948'),('2017-3','D-4176'),('2017-3','D-4294-01'),('2017-3','D-445'),('2017-3','D-4952'),('2017-3','D-56'),('2017-3','D-86-01'),('2017-3','D-86-03'),('2017-3','D-86-07'),('2017-3','D-86-11'),('2017-3','D-86-24'),('2017-3','D-86-25'),('2017-3','D-86-26'),('2017-30','D-525'),('2017-31','D-381'),('2017-32','D-381'),('2017-33','D-3338'),('2017-34','D-1250'),('2017-34','D-1298'),('2017-34','D-2500'),('2017-35','D-1094-01'),('2017-35','D-1094-02'),('2017-35','D-1094-03'),('2017-35','D-1250'),('2017-35','D-1298'),('2017-35','D-130-01'),('2017-35','D-130-02'),('2017-35','D-1319-01'),('2017-35','D-1319-02'),('2017-35','D-1319-03'),('2017-35','D-1322'),('2017-35','D-1840'),('2017-35','D-2386'),('2017-35','D-2500'),('2017-35','D-323'),('2017-35','D-3241-01'),('2017-35','D-3241-02'),('2017-35','D-3242'),('2017-35','D-3338'),('2017-35','D-381'),('2017-35','D-3948'),('2017-35','D-4176'),('2017-35','D-4294-01'),('2017-35','D-4294-02'),('2017-35','D-445'),('2017-35','D-4815'),('2017-35','D-4952'),('2017-35','D-4953'),('2017-35','D-525'),('2017-35','D-56'),('2017-35','D-7321'),('2017-35','D-7978-01'),('2017-35','D-7978-02'),('2017-35','D-86-01'),('2017-35','D-86-02'),('2017-35','D-86-03'),('2017-35','D-86-04'),('2017-35','D-86-05'),('2017-35','D-86-06'),('2017-35','D-86-07'),('2017-35','D-86-08'),('2017-35','D-86-09'),('2017-35','D-86-10'),('2017-35','D-86-11'),('2017-35','D-86-12'),('2017-35','D-86-13'),('2017-35','D-86-14'),('2017-35','D-86-15'),('2017-35','D-86-16'),('2017-35','D-86-17'),('2017-35','D-86-18'),('2017-35','D-86-19'),('2017-35','D-86-20'),('2017-35','D-86-21'),('2017-35','D-86-22'),('2017-35','D-86-23'),('2017-35','D-86-24'),('2017-35','D-86-25'),('2017-35','D-86-26'),('2017-35','D-86-27'),('2017-35','D-86-28'),('2017-35','D-93'),('2017-35','D-97'),('2017-35','DM-01'),('2017-35','DM-02'),('2017-35','E-1064'),('2017-35','EN12937'),('2017-35','EN14078'),('2017-35','EN14103-01'),('2017-35','EN14103-02'),('2017-35','HM-01'),('2017-35','HM-02'),('2017-35','MMIN'),('2017-35','VISUAL'),('2017-4','D-1094-01'),('2017-4','D-1094-02'),('2017-4','D-1094-03'),('2017-4','D-1250'),('2017-4','D-1298'),('2017-4','D-130-01'),('2017-4','D-2386'),('2017-4','D-381'),('2017-4','D-3948'),('2017-4','D-4176'),('2017-4','D-56'),('2017-4','D-86-01'),('2017-4','D-86-03'),('2017-4','D-86-07'),('2017-4','D-86-11'),('2017-4','D-86-24'),('2017-4','D-86-25'),('2017-4','D-86-26'),('2017-6','D-130-02'),('2017-6','D-1319-01'),('2017-6','D-1319-02'),('2017-6','D-1319-03'),('2017-6','D-381'),('2017-6','D-4953'),('2017-6','D-525'),('2017-6','D-86-01'),('2017-6','D-86-14'),('2017-6','D-86-18'),('2017-6','D-86-22'),('2017-6','D-86-24'),('2017-6','D-86-28'),('2017-7','D-1094-01'),('2017-7','D-1094-02'),('2017-7','D-1094-03'),('2017-7','D-1250'),('2017-7','D-1298'),('2017-7','D-130-01'),('2017-7','D-2386'),('2017-7','D-381'),('2017-7','D-3948'),('2017-7','D-4176'),('2017-7','D-56'),('2017-7','D-86-01'),('2017-7','D-86-03'),('2017-7','D-86-07'),('2017-7','D-86-11'),('2017-7','D-86-24'),('2017-7','D-86-25'),('2017-7','D-86-26'),('2017-8','D-1094-01'),('2017-8','D-1094-02'),('2017-8','D-1094-03'),('2017-8','D-1250'),('2017-8','D-1298'),('2017-8','D-130-01'),('2017-8','D-1319-01'),('2017-8','D-1322'),('2017-8','D-1840'),('2017-8','D-2386'),('2017-8','D-3241-01'),('2017-8','D-3241-02'),('2017-8','D-3242'),('2017-8','D-3338'),('2017-8','D-381'),('2017-8','D-3948'),('2017-8','D-4176'),('2017-8','D-4294-01'),('2017-8','D-445'),('2017-8','D-4952'),('2017-8','D-56'),('2017-8','D-86-01'),('2017-8','D-86-03'),('2017-8','D-86-07'),('2017-8','D-86-11'),('2017-8','D-86-24'),('2017-8','D-86-25'),('2017-8','D-86-26'),('2017-9','D-3242'),('R-2017-3a','D-1094-01'),('R-2017-3a','D-1094-02'),('R-2017-3a','D-1094-03'),('R-2017-3a','D-1250'),('R-2017-3a','D-1298'),('R-2017-3a','D-130-01'),('R-2017-3a','D-1319-01'),('R-2017-3a','D-1322'),('R-2017-3a','D-1840'),('R-2017-3a','D-2386'),('R-2017-3a','D-3241-01'),('R-2017-3a','D-3241-02'),('R-2017-3a','D-3242'),('R-2017-3a','D-3338'),('R-2017-3a','D-381'),('R-2017-3a','D-3948'),('R-2017-3a','D-4176'),('R-2017-3a','D-4294-01'),('R-2017-3a','D-445'),('R-2017-3a','D-4952'),('R-2017-3a','D-56'),('R-2017-3a','D-86-01'),('R-2017-3a','D-86-03'),('R-2017-3a','D-86-07'),('R-2017-3a','D-86-11'),('R-2017-3a','D-86-24'),('R-2017-3a','D-86-25'),('R-2017-3a','D-86-26');
/*!40000 ALTER TABLE `pruebasxproducto` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `reportes`
--

DROP TABLE IF EXISTS `reportes`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `reportes` (
  `ProductoID` varchar(20) NOT NULL COMMENT 'ID del producto',
  `ClienteID` smallint(6) NOT NULL COMMENT 'Id del Cliente',
  `Tipo_Producto_ID` smallint(6) NOT NULL COMMENT 'Llave Primaria',
  `Fecha_Ingreso` date NOT NULL COMMENT 'Fecha de ingreso del producto proveniente de la tabla productos',
  `Hora_Ingreso` time NOT NULL,
  `Fecha_Reporte` date NOT NULL COMMENT 'Fecha en la que se genero el reporte',
  `ID_Prueba` varchar(20) NOT NULL COMMENT 'Prueba que se realizo',
  `Valor` varchar(100) NOT NULL COMMENT 'Valor obtenido para la respectica subprueba',
  `Valor_2` varchar(100) DEFAULT NULL,
  `Valor_3` varchar(100) DEFAULT NULL,
  `Valor_4` varchar(100) DEFAULT NULL,
  `Valor_5` varchar(100) DEFAULT NULL,
  `UsuarioID` smallint(6) NOT NULL COMMENT 'ID del usuario que llevo a cabo el reporte',
  PRIMARY KEY (`ProductoID`,`ID_Prueba`),
  KEY `FK_ReportesPruebas` (`ID_Prueba`),
  KEY `FK_ReportesUsuarios` (`UsuarioID`),
  KEY `FK_ReportesClientes` (`ClienteID`),
  KEY `FK_ReportesTipoProductoID` (`Tipo_Producto_ID`),
  CONSTRAINT `reportes_ibfk_1` FOREIGN KEY (`ProductoID`) REFERENCES `productos` (`ProductoID`),
  CONSTRAINT `reportes_ibfk_2` FOREIGN KEY (`ID_Prueba`) REFERENCES `pruebas` (`ID_Prueba`),
  CONSTRAINT `reportes_ibfk_3` FOREIGN KEY (`UsuarioID`) REFERENCES `usuarios` (`UsuarioID`),
  CONSTRAINT `reportes_ibfk_4` FOREIGN KEY (`ClienteID`) REFERENCES `clientes` (`ClienteID`),
  CONSTRAINT `reportes_ibfk_5` FOREIGN KEY (`Tipo_Producto_ID`) REFERENCES `tipo_productos` (`Tipo_Producto_ID`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `reportes`
--

LOCK TABLES `reportes` WRITE;
/*!40000 ALTER TABLE `reportes` DISABLE KEYS */;
INSERT INTO `reportes` VALUES ('2017-11',1,4,'2017-11-10','08:26:50','2017-11-10','D-1319-01','123',NULL,NULL,NULL,NULL,5),('2017-11',1,4,'2017-11-10','08:26:50','2017-11-10','D-1322','45',NULL,NULL,NULL,NULL,5),('2017-11',1,4,'2017-11-10','08:26:50','2017-11-10','D-1840','65',NULL,NULL,NULL,NULL,5),('2017-12',1,14,'2017-11-10','09:13:54','2017-11-10','D-1319-01','18,6',NULL,NULL,NULL,NULL,5),('2017-12',1,14,'2017-11-10','09:13:54','2017-11-10','D-1319-02','1,5',NULL,NULL,NULL,NULL,5),('2017-12',1,14,'2017-11-10','09:13:54','2017-11-10','D-1319-03','80,0',NULL,NULL,NULL,NULL,5),('2017-12',1,14,'2017-11-10','09:13:54','2017-11-10','D-1322','19,5','19,5',NULL,NULL,NULL,5),('2017-12',1,14,'2017-11-10','09:13:54','2017-11-10','D-1840','1,01',NULL,NULL,NULL,NULL,5),('2017-12',1,14,'2017-11-10','09:13:54','2017-11-10','D-3242','0,053',NULL,NULL,NULL,NULL,5),('2017-13',1,14,'2017-11-10','13:19:31','2017-11-10','D-56','44',NULL,NULL,NULL,NULL,5),('2017-13',1,14,'2017-11-10','13:19:31','2017-11-10','D-93','46',NULL,NULL,NULL,NULL,5),('2017-14',1,14,'2017-11-10','15:12:08','2017-11-10','D-4294-01','13,9','13,2',NULL,NULL,NULL,5),('2017-14',1,14,'2017-11-10','15:12:08','2017-11-10','D-4294-02','9,55','10,5',NULL,NULL,NULL,5),('2017-15',8,14,'2017-11-14','08:47:45','2017-11-14','D-3241-01','Caida Presion',NULL,NULL,NULL,NULL,5),('2017-15',8,14,'2017-11-14','08:47:45','2017-11-14','D-3241-02','Rojo',NULL,NULL,NULL,NULL,5),('2017-16',2,14,'2017-11-14','09:42:29','2017-11-14','D-86-01','100',NULL,NULL,NULL,NULL,5),('2017-16',2,14,'2017-11-14','09:42:29','2017-11-14','D-86-03','110',NULL,NULL,NULL,NULL,5),('2017-16',2,14,'2017-11-14','09:42:29','2017-11-14','D-86-07','150',NULL,NULL,NULL,NULL,5),('2017-16',2,14,'2017-11-14','09:42:29','2017-11-14','D-86-11','190',NULL,NULL,NULL,NULL,5),('2017-16',2,14,'2017-11-14','09:42:29','2017-11-14','D-86-12','195',NULL,NULL,NULL,NULL,5),('2017-16',2,14,'2017-11-14','09:42:29','2017-11-14','D-86-14','110',NULL,NULL,NULL,NULL,5),('2017-16',2,14,'2017-11-14','09:42:29','2017-11-14','D-86-17','140',NULL,NULL,NULL,NULL,5),('2017-16',2,14,'2017-11-14','09:42:29','2017-11-14','D-86-18','150',NULL,NULL,NULL,NULL,5),('2017-16',2,14,'2017-11-14','09:42:29','2017-11-14','D-86-22','190',NULL,NULL,NULL,NULL,5),('2017-16',2,14,'2017-11-14','09:42:29','2017-11-14','D-86-24','200',NULL,NULL,NULL,NULL,5),('2017-16',2,14,'2017-11-14','09:42:29','2017-11-14','D-86-25','20',NULL,NULL,NULL,NULL,5),('2017-16',2,14,'2017-11-14','09:42:29','2017-11-14','D-86-26','10',NULL,NULL,NULL,NULL,5),('2017-16',2,14,'2017-11-14','09:42:29','2017-11-14','D-86-27','80',NULL,NULL,NULL,NULL,5),('2017-16',2,14,'2017-11-14','09:42:29','2017-11-14','D-86-28','20',NULL,NULL,NULL,NULL,5),('2017-17',1,14,'2017-11-14','10:27:24','2017-11-14','D-2386','-45,5',NULL,NULL,NULL,NULL,5),('2017-18',1,14,'2017-11-14','10:53:49','2017-11-14','D-4176','NO PASA',NULL,NULL,NULL,NULL,2),('2017-19',1,14,'2017-11-14','12:02:26','2017-11-14','D-1094-01','40',NULL,NULL,NULL,NULL,5),('2017-19',1,14,'2017-11-14','12:02:26','2017-11-14','D-1094-02','2',NULL,NULL,NULL,NULL,5),('2017-19',1,14,'2017-11-14','12:02:26','2017-11-14','D-1094-03','1',NULL,NULL,NULL,NULL,5),('2017-19',1,14,'2017-11-14','12:02:26','2017-11-14','D-3948','35',NULL,NULL,NULL,NULL,5),('2017-20',1,14,'2017-11-14','13:18:38','2017-11-14','D-130-01','50',NULL,NULL,NULL,NULL,5),('2017-20',1,14,'2017-11-14','13:18:38','2017-11-14','D-130-02','25',NULL,NULL,NULL,NULL,5),('2017-21',1,14,'2017-11-14','14:03:47','2017-11-14','D-4953','565,50',NULL,NULL,NULL,NULL,5),('2017-22',1,14,'2017-11-14','14:34:47','2017-11-14','D-4815','60',NULL,NULL,NULL,NULL,5),('2017-23',1,14,'2017-11-14','15:03:36','2017-11-14','D-1298','25',NULL,NULL,NULL,NULL,5),('2017-24',1,14,'2017-11-14','15:27:40','2017-11-14','D-1250','20',NULL,NULL,NULL,NULL,5),('2017-24',1,14,'2017-11-14','15:27:40','2017-11-14','D-4952','POSITIVO',NULL,NULL,NULL,NULL,5),('2017-24',1,14,'2017-11-14','15:27:40','2017-11-14','E-1064','65',NULL,NULL,NULL,NULL,5),('2017-25',1,14,'2017-11-14','15:44:50','2017-11-14','D-445','10',NULL,NULL,NULL,NULL,5),('2017-26',1,14,'2017-11-15','08:19:11','2017-11-15','MMIN','610,7',NULL,NULL,NULL,NULL,5),('2017-28',1,14,'2017-11-15','09:04:32','2017-11-15','EN14078','65',NULL,NULL,NULL,NULL,5),('2017-29',1,14,'2017-11-15','10:24:37','2017-11-15','D-7321','1500000',NULL,NULL,NULL,NULL,5),('2017-29',1,14,'2017-11-15','10:24:37','2017-11-15','EN12937','82,974',NULL,NULL,NULL,NULL,5),('2017-3',1,4,'2017-11-02','10:13:19','2017-11-02','D-1094-01','123',NULL,NULL,NULL,NULL,5),('2017-3',1,4,'2017-11-02','10:13:19','2017-11-02','D-1094-02','123',NULL,NULL,NULL,NULL,5),('2017-3',1,4,'2017-11-02','10:13:19','2017-11-02','D-1094-03','123',NULL,NULL,NULL,NULL,5),('2017-3',1,4,'2017-11-02','10:13:19','2017-11-02','D-1250','123',NULL,NULL,NULL,NULL,5),('2017-3',1,4,'2017-11-02','10:13:19','2017-11-02','D-1298','123',NULL,NULL,NULL,NULL,5),('2017-3',1,4,'2017-11-02','10:13:19','2017-11-02','D-130-01','123',NULL,NULL,NULL,NULL,5),('2017-3',1,4,'2017-11-02','10:13:19','2017-11-02','D-1319-01','123',NULL,NULL,NULL,NULL,5),('2017-3',1,4,'2017-11-02','10:13:19','2017-11-02','D-1322','123',NULL,NULL,NULL,NULL,5),('2017-3',1,4,'2017-11-02','10:13:19','2017-11-02','D-1840','123',NULL,NULL,NULL,NULL,5),('2017-3',1,4,'2017-11-02','10:13:19','2017-11-02','D-2386','123',NULL,NULL,NULL,NULL,5),('2017-3',1,4,'2017-11-02','10:13:19','2017-11-02','D-3241-01','123',NULL,NULL,NULL,NULL,5),('2017-3',1,4,'2017-11-02','10:13:19','2017-11-02','D-3241-02','123',NULL,NULL,NULL,NULL,5),('2017-3',1,4,'2017-11-02','10:13:19','2017-11-02','D-3242','123',NULL,NULL,NULL,NULL,5),('2017-3',1,4,'2017-11-02','10:13:19','2017-11-02','D-3338','123',NULL,NULL,NULL,NULL,5),('2017-3',1,4,'2017-11-02','10:13:19','2017-11-02','D-381','123',NULL,NULL,NULL,NULL,5),('2017-3',1,4,'2017-11-02','10:13:19','2017-11-02','D-3948','123',NULL,NULL,NULL,NULL,5),('2017-3',1,4,'2017-11-02','10:13:19','2017-11-02','D-4176','123',NULL,NULL,NULL,NULL,5),('2017-3',1,4,'2017-11-02','10:13:19','2017-11-02','D-4294-01','123',NULL,NULL,NULL,NULL,5),('2017-3',1,4,'2017-11-02','10:13:19','2017-11-02','D-445','123',NULL,NULL,NULL,NULL,5),('2017-3',1,4,'2017-11-02','10:13:19','2017-11-02','D-4952','123',NULL,NULL,NULL,NULL,5),('2017-3',1,4,'2017-11-02','10:13:19','2017-11-02','D-56','123',NULL,NULL,NULL,NULL,5),('2017-3',1,4,'2017-11-02','10:13:19','2017-11-02','D-86-01','123',NULL,NULL,NULL,NULL,5),('2017-3',1,4,'2017-11-02','10:13:19','2017-11-02','D-86-03','123',NULL,NULL,NULL,NULL,5),('2017-3',1,4,'2017-11-02','10:13:19','2017-11-02','D-86-07','123',NULL,NULL,NULL,NULL,5),('2017-3',1,4,'2017-11-02','10:13:19','2017-11-02','D-86-11','123',NULL,NULL,NULL,NULL,5),('2017-3',1,4,'2017-11-02','10:13:19','2017-11-02','D-86-24','123',NULL,NULL,NULL,NULL,5),('2017-3',1,4,'2017-11-02','10:13:19','2017-11-02','D-86-25','123',NULL,NULL,NULL,NULL,5),('2017-3',1,4,'2017-11-02','10:13:19','2017-11-02','D-86-26','123',NULL,NULL,NULL,NULL,5),('2017-30',1,14,'2017-11-15','10:56:42','2017-11-15','D-525','11,01',NULL,NULL,NULL,NULL,5),('2017-31',1,4,'2017-11-15','13:20:54','2017-11-15','D-381','<1',NULL,NULL,NULL,NULL,5),('2017-32',1,6,'2017-11-15','13:21:14','2017-11-15','D-381','<0.5',NULL,NULL,NULL,NULL,5),('2017-33',1,14,'2017-11-15','14:26:45','2017-11-15','D-3338','298,917',NULL,NULL,NULL,NULL,5),('2017-4',6,3,'2017-11-02','10:14:06','2017-11-02','D-1094-01','123',NULL,NULL,NULL,NULL,2),('2017-4',6,3,'2017-11-02','10:14:06','2017-11-02','D-1094-02','v',NULL,NULL,NULL,NULL,2),('2017-4',6,3,'2017-11-02','10:14:06','2017-11-02','D-1094-03','123',NULL,NULL,NULL,NULL,2),('2017-4',6,3,'2017-11-02','10:14:06','2017-11-02','D-1250','123',NULL,NULL,NULL,NULL,2),('2017-4',6,3,'2017-11-02','10:14:06','2017-11-02','D-1298','123',NULL,NULL,NULL,NULL,2),('2017-4',6,3,'2017-11-02','10:14:06','2017-11-02','D-130-01','123',NULL,NULL,NULL,NULL,2),('2017-4',6,3,'2017-11-02','10:14:06','2017-11-02','D-2386','123',NULL,NULL,NULL,NULL,2),('2017-4',6,3,'2017-11-02','10:14:06','2017-11-02','D-381','123',NULL,NULL,NULL,NULL,2),('2017-4',6,3,'2017-11-02','10:14:06','2017-11-02','D-3948','123',NULL,NULL,NULL,NULL,2),('2017-4',6,3,'2017-11-02','10:14:06','2017-11-02','D-4176','123',NULL,NULL,NULL,NULL,2),('2017-4',6,3,'2017-11-02','10:14:06','2017-11-02','D-56','123',NULL,NULL,NULL,NULL,2),('2017-4',6,3,'2017-11-02','10:14:06','2017-11-02','D-86-01','123',NULL,NULL,NULL,NULL,2),('2017-4',6,3,'2017-11-02','10:14:06','2017-11-02','D-86-03','123',NULL,NULL,NULL,NULL,2),('2017-4',6,3,'2017-11-02','10:14:06','2017-11-02','D-86-07','123',NULL,NULL,NULL,NULL,2),('2017-4',6,3,'2017-11-02','10:14:06','2017-11-02','D-86-11','123',NULL,NULL,NULL,NULL,2),('2017-4',6,3,'2017-11-02','10:14:06','2017-11-02','D-86-24','123',NULL,NULL,NULL,NULL,2),('2017-4',6,3,'2017-11-02','10:14:06','2017-11-02','D-86-25','123',NULL,NULL,NULL,NULL,2),('2017-4',6,3,'2017-11-02','10:14:06','2017-11-02','D-86-26','123',NULL,NULL,NULL,NULL,2),('2017-6',1,6,'2017-11-03','10:20:14','2017-11-03','D-130-02','123',NULL,NULL,NULL,NULL,2),('2017-6',1,6,'2017-11-03','10:20:14','2017-11-03','D-1319-01','12312',NULL,NULL,NULL,NULL,2),('2017-6',1,6,'2017-11-03','10:20:14','2017-11-03','D-1319-02','123321',NULL,NULL,NULL,NULL,2),('2017-6',1,6,'2017-11-03','10:20:14','2017-11-03','D-1319-03','321',NULL,NULL,NULL,NULL,2),('2017-6',1,6,'2017-11-03','10:20:14','2017-11-03','D-381','1231',NULL,NULL,NULL,NULL,2),('2017-6',1,6,'2017-11-03','10:20:14','2017-11-03','D-4953','12312',NULL,NULL,NULL,NULL,2),('2017-6',1,6,'2017-11-03','10:20:14','2017-11-03','D-525','32131',NULL,NULL,NULL,NULL,2),('2017-6',1,6,'2017-11-03','10:20:14','2017-11-03','D-86-01','321312',NULL,NULL,NULL,NULL,2),('2017-6',1,6,'2017-11-03','10:20:14','2017-11-03','D-86-14','123123',NULL,NULL,NULL,NULL,2),('2017-6',1,6,'2017-11-03','10:20:14','2017-11-03','D-86-18','312312',NULL,NULL,NULL,NULL,2),('2017-6',1,6,'2017-11-03','10:20:14','2017-11-03','D-86-22','123123123',NULL,NULL,NULL,NULL,2),('2017-6',1,6,'2017-11-03','10:20:14','2017-11-03','D-86-24','1231231',NULL,NULL,NULL,NULL,2),('2017-6',1,6,'2017-11-03','10:20:14','2017-11-03','D-86-28','12312',NULL,NULL,NULL,NULL,2),('2017-7',1,4,'2017-11-03','15:48:59','2017-11-03','D-1094-01','123',NULL,NULL,NULL,NULL,5),('2017-7',1,4,'2017-11-03','15:48:59','2017-11-03','D-1094-02','123',NULL,NULL,NULL,NULL,5),('2017-7',1,4,'2017-11-03','15:48:59','2017-11-03','D-1094-03','123',NULL,NULL,NULL,NULL,5),('2017-7',1,4,'2017-11-03','15:48:59','2017-11-03','D-1250','123',NULL,NULL,NULL,NULL,5),('2017-7',1,4,'2017-11-03','15:48:59','2017-11-03','D-1298','123',NULL,NULL,NULL,NULL,5),('2017-7',1,4,'2017-11-03','15:48:59','2017-11-03','D-130-01','123',NULL,NULL,NULL,NULL,5),('2017-7',1,4,'2017-11-03','15:48:59','2017-11-03','D-2386','123',NULL,NULL,NULL,NULL,5),('2017-7',1,4,'2017-11-03','15:48:59','2017-11-03','D-381','123',NULL,NULL,NULL,NULL,5),('2017-7',1,4,'2017-11-03','15:48:59','2017-11-03','D-3948','123',NULL,NULL,NULL,NULL,5),('2017-7',1,4,'2017-11-03','15:48:59','2017-11-03','D-4176','123',NULL,NULL,NULL,NULL,5),('2017-7',1,4,'2017-11-03','15:48:59','2017-11-03','D-56','123',NULL,NULL,NULL,NULL,5),('2017-7',1,4,'2017-11-03','15:48:59','2017-11-03','D-86-01','123',NULL,NULL,NULL,NULL,5),('2017-7',1,4,'2017-11-03','15:48:59','2017-11-03','D-86-03','123',NULL,NULL,NULL,NULL,5),('2017-7',1,4,'2017-11-03','15:48:59','2017-11-03','D-86-07','123',NULL,NULL,NULL,NULL,5),('2017-7',1,4,'2017-11-03','15:48:59','2017-11-03','D-86-11','123',NULL,NULL,NULL,NULL,5),('2017-7',1,4,'2017-11-03','15:48:59','2017-11-03','D-86-24','123',NULL,NULL,NULL,NULL,5),('2017-7',1,4,'2017-11-03','15:48:59','2017-11-03','D-86-25','123',NULL,NULL,NULL,NULL,5),('2017-7',1,4,'2017-11-03','15:48:59','2017-11-03','D-86-26','123',NULL,NULL,NULL,NULL,5),('2017-9',1,4,'2017-11-08','10:15:19','2017-11-08','D-3242','0,053',NULL,NULL,NULL,NULL,5),('R-2017-3a',1,4,'2017-11-15','15:56:34','2017-11-15','D-1094-01','40',NULL,NULL,NULL,NULL,5),('R-2017-3a',1,4,'2017-11-15','15:56:34','2017-11-15','D-1094-02','2',NULL,NULL,NULL,NULL,5),('R-2017-3a',1,4,'2017-11-15','15:56:34','2017-11-15','D-1094-03','1',NULL,NULL,NULL,NULL,5),('R-2017-3a',1,4,'2017-11-15','15:56:34','2017-11-15','D-1250','20',NULL,NULL,NULL,NULL,5),('R-2017-3a',1,4,'2017-11-15','15:56:34','2017-11-15','D-1298','25',NULL,NULL,NULL,NULL,5),('R-2017-3a',1,4,'2017-11-15','15:56:34','2017-11-15','D-130-01','50',NULL,NULL,NULL,NULL,5),('R-2017-3a',1,4,'2017-11-15','15:56:34','2017-11-15','D-1319-01','18,6',NULL,NULL,NULL,NULL,5),('R-2017-3a',1,4,'2017-11-15','15:56:34','2017-11-15','D-1322','19,5',NULL,NULL,NULL,NULL,5),('R-2017-3a',1,4,'2017-11-15','15:56:34','2017-11-15','D-1840','1,01',NULL,NULL,NULL,NULL,5),('R-2017-3a',1,4,'2017-11-15','15:56:34','2017-11-15','D-2386','-45,0',NULL,NULL,NULL,NULL,5),('R-2017-3a',1,4,'2017-11-15','15:56:34','2017-11-15','D-3241-01','Caida Presion',NULL,NULL,NULL,NULL,5),('R-2017-3a',1,4,'2017-11-15','15:56:34','2017-11-15','D-3241-02','Rojo',NULL,NULL,NULL,NULL,5),('R-2017-3a',1,4,'2017-11-15','15:56:34','2017-11-15','D-3242','0,053',NULL,NULL,NULL,NULL,5),('R-2017-3a',1,4,'2017-11-15','15:56:34','2017-11-15','D-3338','298,917',NULL,NULL,NULL,NULL,5),('R-2017-3a',1,4,'2017-11-15','15:56:34','2017-11-15','D-381','<1',NULL,NULL,NULL,NULL,5),('R-2017-3a',1,4,'2017-11-15','15:56:34','2017-11-15','D-3948','35',NULL,NULL,NULL,NULL,5),('R-2017-3a',1,4,'2017-11-15','15:56:34','2017-11-15','D-4176','NO PASA',NULL,NULL,NULL,NULL,5),('R-2017-3a',1,4,'2017-11-15','15:56:34','2017-11-15','D-4294-01','13,9',NULL,NULL,NULL,NULL,5),('R-2017-3a',1,4,'2017-11-15','15:56:34','2017-11-15','D-445','10',NULL,NULL,NULL,NULL,5),('R-2017-3a',1,4,'2017-11-15','15:56:34','2017-11-15','D-4952','POSITIVO',NULL,NULL,NULL,NULL,5),('R-2017-3a',1,4,'2017-11-15','15:56:34','2017-11-15','D-56','44',NULL,NULL,NULL,NULL,5),('R-2017-3a',1,4,'2017-11-15','15:56:34','2017-11-15','D-86-01','100',NULL,NULL,NULL,NULL,5),('R-2017-3a',1,4,'2017-11-15','15:56:34','2017-11-15','D-86-03','110',NULL,NULL,NULL,NULL,5),('R-2017-3a',1,4,'2017-11-15','15:56:34','2017-11-15','D-86-07','150',NULL,NULL,NULL,NULL,5),('R-2017-3a',1,4,'2017-11-15','15:56:34','2017-11-15','D-86-11','190',NULL,NULL,NULL,NULL,5),('R-2017-3a',1,4,'2017-11-15','15:56:34','2017-11-15','D-86-24','200',NULL,NULL,NULL,NULL,5),('R-2017-3a',1,4,'2017-11-15','15:56:34','2017-11-15','D-86-25','20',NULL,NULL,NULL,NULL,5),('R-2017-3a',1,4,'2017-11-15','15:56:34','2017-11-15','D-86-26','10',NULL,NULL,NULL,NULL,5);
/*!40000 ALTER TABLE `reportes` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `tipo_productos`
--

DROP TABLE IF EXISTS `tipo_productos`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `tipo_productos` (
  `Tipo_Producto_ID` smallint(6) NOT NULL COMMENT 'Llave Primaria',
  `Nombre` varchar(30) NOT NULL COMMENT 'Nombre del tipo de producto',
  `Intervalo` varchar(25) NOT NULL,
  PRIMARY KEY (`Tipo_Producto_ID`),
  UNIQUE KEY `Nombre_UNIQUE` (`Nombre`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tipo_productos`
--

LOCK TABLES `tipo_productos` WRITE;
/*!40000 ALTER TABLE `tipo_productos` DISABLE KEYS */;
INSERT INTO `tipo_productos` VALUES (1,'ETANOL','8 DAY'),(2,'BIODIESEL (B100)','8 DAY'),(3,'BX','12 DAY'),(4,'JETA-1','3 DAY'),(5,'DIESEL','10 DAY'),(6,'GASOLINA CORRIENTE','8 DAY'),(7,'GASOLINA EXTRA','8 DAY'),(8,'GASOLINA EXTRA OXIGENADA','25 DAY'),(9,'GASOLINA CORRIENTE OXIGENADA','25 DAY'),(10,'HIDROMETRO','25 DAY'),(11,'DENSIMETRO DIGITAL','12 DAY'),(12,'TERMOMETRO TL1A','12 DAY'),(13,'TERMOMETRO TP7','12 DAY'),(14,'Otros','12 DAY');
/*!40000 ALTER TABLE `tipo_productos` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `usuarios`
--

DROP TABLE IF EXISTS `usuarios`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `usuarios` (
  `UsuarioID` smallint(6) NOT NULL COMMENT 'Id del analista',
  `usuario` varchar(20) NOT NULL COMMENT 'Nombre del analista',
  `Contraseña` varchar(20) NOT NULL DEFAULT '123' COMMENT 'Contraseña del analista',
  `Admin` tinyint(1) NOT NULL DEFAULT '0',
  PRIMARY KEY (`UsuarioID`),
  UNIQUE KEY `usuario_UNIQUE` (`usuario`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `usuarios`
--

LOCK TABLES `usuarios` WRITE;
/*!40000 ALTER TABLE `usuarios` DISABLE KEYS */;
INSERT INTO `usuarios` VALUES (1,'Administrador','Admin1513',1),(2,'Israel Olarte','123',1),(3,'Carlos Sandoval','123',0),(4,'Analista 3','123',0),(5,'Nicolas','123',1);
/*!40000 ALTER TABLE `usuarios` ENABLE KEYS */;
UNLOCK TABLES;
/*!40103 SET TIME_ZONE=@OLD_TIME_ZONE */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;

-- Dump completed on 2018-01-09 15:45:12
