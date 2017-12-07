-- MySQL dump 10.13  Distrib 5.7.17, for Win64 (x86_64)
--
-- Host: localhost    Database: bd_inventario
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
-- Table structure for table `categorias`
--

DROP TABLE IF EXISTS `categorias`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `categorias` (
  `Id_Categoria` int(11) NOT NULL AUTO_INCREMENT,
  `Nombre_Categoria` varchar(45) DEFAULT NULL,
  PRIMARY KEY (`Id_Categoria`)
) ENGINE=InnoDB AUTO_INCREMENT=8 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `categorias`
--

LOCK TABLES `categorias` WRITE;
/*!40000 ALTER TABLE `categorias` DISABLE KEYS */;
INSERT INTO `categorias` VALUES (1,'Prueba categoria 1'),(2,'Prueba categoria 2'),(3,'pruebaT0'),(4,'catnew1p'),(6,'prueba'),(7,'pruebaCatH');
/*!40000 ALTER TABLE `categorias` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `categorias_sub`
--

DROP TABLE IF EXISTS `categorias_sub`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `categorias_sub` (
  `Id_SubCategoria` int(11) NOT NULL AUTO_INCREMENT,
  `Nombre_SubCategoria` varchar(45) DEFAULT NULL,
  `Id_Categoria` int(11) DEFAULT NULL,
  PRIMARY KEY (`Id_SubCategoria`),
  KEY `fk_Categorias_Sub_Categorias_idx` (`Id_Categoria`),
  CONSTRAINT `fk_Categorias_Sub_Categorias` FOREIGN KEY (`Id_Categoria`) REFERENCES `categorias` (`Id_Categoria`) ON DELETE NO ACTION ON UPDATE NO ACTION
) ENGINE=InnoDB AUTO_INCREMENT=19 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `categorias_sub`
--

LOCK TABLES `categorias_sub` WRITE;
/*!40000 ALTER TABLE `categorias_sub` DISABLE KEYS */;
INSERT INTO `categorias_sub` VALUES (1,'test01',1),(2,'test02',1),(3,'test3',2),(4,'test4',2),(5,'test5',3),(6,'test6',3),(7,'Subcatnew01',4),(8,'test04.',1),(10,'sub1',6),(11,'sub2',6),(12,'sub3',6),(16,'sub4',6),(17,'subcat pruebat',3),(18,'nuevasubPrueba',7);
/*!40000 ALTER TABLE `categorias_sub` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `cotizacion_solicitud`
--

DROP TABLE IF EXISTS `cotizacion_solicitud`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `cotizacion_solicitud` (
  `IdCotizacion_Solicitud` int(11) NOT NULL AUTO_INCREMENT,
  `IdSolicitud_Compras` int(11) DEFAULT NULL,
  `Cotizacion` varchar(200) DEFAULT NULL,
  PRIMARY KEY (`IdCotizacion_Solicitud`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `cotizacion_solicitud`
--

LOCK TABLES `cotizacion_solicitud` WRITE;
/*!40000 ALTER TABLE `cotizacion_solicitud` DISABLE KEYS */;
/*!40000 ALTER TABLE `cotizacion_solicitud` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `datos_app`
--

DROP TABLE IF EXISTS `datos_app`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `datos_app` (
  `IdDatos_App` int(11) NOT NULL AUTO_INCREMENT,
  `Nombre` varchar(45) DEFAULT NULL,
  `Detalles` varchar(200) DEFAULT NULL,
  PRIMARY KEY (`IdDatos_App`)
) ENGINE=InnoDB AUTO_INCREMENT=8 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `datos_app`
--

LOCK TABLES `datos_app` WRITE;
/*!40000 ALTER TABLE `datos_app` DISABLE KEYS */;
INSERT INTO `datos_app` VALUES (1,'Estante','3'),(2,'Entrepano','4'),(3,'Caja','4'),(4,'Azul','Salud'),(5,'Rojo','Inflamabilidad'),(6,'Amarillo','Reactividad'),(7,'Blanco','Especifico');
/*!40000 ALTER TABLE `datos_app` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `doag`
--

DROP TABLE IF EXISTS `doag`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `doag` (
  `Id_Doag` int(11) NOT NULL AUTO_INCREMENT,
  `Nombre_Doag` varchar(45) DEFAULT NULL,
  `Monto` double DEFAULT NULL,
  `Comentario` varchar(45) DEFAULT NULL,
  PRIMARY KEY (`Id_Doag`)
) ENGINE=InnoDB AUTO_INCREMENT=6 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `doag`
--

LOCK TABLES `doag` WRITE;
/*!40000 ALTER TABLE `doag` DISABLE KEYS */;
INSERT INTO `doag` VALUES (1,'Basico',2500000,NULL),(2,'Certificado',7000000,NULL),(3,'Advanced',25000000,NULL),(4,'Elite',100000000,'t'),(5,'Prueba3',9000099,'otro');
/*!40000 ALTER TABLE `doag` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `equipos`
--

DROP TABLE IF EXISTS `equipos`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `equipos` (
  `Id_Equipo` int(11) NOT NULL AUTO_INCREMENT,
  `Cod_Equipo` varchar(45) DEFAULT NULL,
  `Nombre_Equipo` varchar(65) DEFAULT NULL,
  `Marca` varchar(45) DEFAULT NULL,
  `Serie` varchar(45) DEFAULT NULL,
  `Foto` text,
  `Activo` tinyint(1) DEFAULT NULL,
  PRIMARY KEY (`Id_Equipo`)
) ENGINE=InnoDB AUTO_INCREMENT=7 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `equipos`
--

LOCK TABLES `equipos` WRITE;
/*!40000 ALTER TABLE `equipos` DISABLE KEYS */;
INSERT INTO `equipos` VALUES (1,'34FGHR76590','INFRARROJO','PRUEBA 0','FGRTY643','C:/Users/Msi/Desktop/Imagenes_Prueba/E_34FGHR76590_1.png',1),(2,'RHT3452346','VISCOSIMETRO','PRUEBA 1..','TEIO4353',NULL,0),(3,'ERGF45373','PRUEBA EQUIPO','PRUEBA 2','56HFD4',NULL,1),(4,'80026568','SILLA','RIMAX','AS14254','C:/Users/Msi/Desktop/Imagenes_Prueba/E_80026568_4.png',1),(5,'GVJH76GHJ','BARRA DE CARBON','GANDO','GH54YG','C:/Users/Msi/Desktop/Imagenes_Prueba/E_GVJH76GHJ_5.png',0),(6,'LP001','DENSIMETRO','TOLEDO','80026','C:/Users/Msi/Desktop/Imagenes_Prueba/E_FP005_6.png',1);
/*!40000 ALTER TABLE `equipos` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `eval_proveedores`
--

DROP TABLE IF EXISTS `eval_proveedores`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `eval_proveedores` (
  `IdEval_Proveedor` int(11) NOT NULL AUTO_INCREMENT,
  `Fecha` date DEFAULT NULL,
  `Nit_Proveedor` varchar(45) DEFAULT NULL,
  `Calidad` int(11) DEFAULT NULL,
  `Servicio` int(11) DEFAULT NULL,
  `Tiempo` int(11) DEFAULT NULL,
  `Costo` int(11) DEFAULT NULL,
  `Soporte` int(11) DEFAULT NULL,
  `Resultado_Eval` varchar(45) DEFAULT NULL,
  PRIMARY KEY (`IdEval_Proveedor`),
  KEY `fk_Eval_Proveedores_Proveedores1_idx` (`Nit_Proveedor`),
  CONSTRAINT `fk_Eval_Proveedores_Proveedores1` FOREIGN KEY (`Nit_Proveedor`) REFERENCES `proveedores` (`Nit_Proveedor`) ON DELETE NO ACTION ON UPDATE NO ACTION
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `eval_proveedores`
--

LOCK TABLES `eval_proveedores` WRITE;
/*!40000 ALTER TABLE `eval_proveedores` DISABLE KEYS */;
/*!40000 ALTER TABLE `eval_proveedores` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `firmas`
--

DROP TABLE IF EXISTS `firmas`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `firmas` (
  `Id_Firma` int(11) NOT NULL AUTO_INCREMENT,
  `Id_Usuario` int(11) DEFAULT NULL,
  `Firma` varchar(45) DEFAULT NULL,
  `Nombre_Doc` varchar(45) DEFAULT NULL,
  PRIMARY KEY (`Id_Firma`),
  KEY `fk_Firmas_Usuarios1_idx` (`Id_Usuario`),
  CONSTRAINT `fk_Firmas_Usuarios1` FOREIGN KEY (`Id_Usuario`) REFERENCES `usuarios` (`Id_Usuario`) ON DELETE NO ACTION ON UPDATE NO ACTION
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `firmas`
--

LOCK TABLES `firmas` WRITE;
/*!40000 ALTER TABLE `firmas` DISABLE KEYS */;
/*!40000 ALTER TABLE `firmas` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `lista_productos`
--

DROP TABLE IF EXISTS `lista_productos`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `lista_productos` (
  `IdLista_Productos` int(11) NOT NULL AUTO_INCREMENT,
  `Id_Producto` int(11) DEFAULT NULL,
  `Cantidad` varchar(45) DEFAULT NULL,
  `IdSolicitud_Compra` int(11) DEFAULT NULL,
  PRIMARY KEY (`IdLista_Productos`),
  KEY `fk_Lista_Productos_Productos1_idx` (`Id_Producto`),
  KEY `fk_Lista_Productos_Solicitud_Compras1_idx` (`IdSolicitud_Compra`),
  CONSTRAINT `fk_Lista_Productos_Productos1` FOREIGN KEY (`Id_Producto`) REFERENCES `productos` (`Id_Producto`) ON DELETE NO ACTION ON UPDATE NO ACTION,
  CONSTRAINT `fk_Lista_Productos_Solicitud_Compras1` FOREIGN KEY (`IdSolicitud_Compra`) REFERENCES `solicitud_compras` (`IdSolicitud_Compra`) ON DELETE NO ACTION ON UPDATE NO ACTION
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `lista_productos`
--

LOCK TABLES `lista_productos` WRITE;
/*!40000 ALTER TABLE `lista_productos` DISABLE KEYS */;
/*!40000 ALTER TABLE `lista_productos` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `lista_proveedores`
--

DROP TABLE IF EXISTS `lista_proveedores`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `lista_proveedores` (
  `IdLista_Proveedor` int(11) NOT NULL AUTO_INCREMENT,
  `IdLista_Producto` int(11) DEFAULT NULL,
  `Precio` varchar(45) DEFAULT NULL,
  `Nit_Proveedor` varchar(45) DEFAULT NULL,
  `IdEval_Proveedor` int(11) DEFAULT NULL,
  `Num_Cotizacion` varchar(45) DEFAULT NULL,
  PRIMARY KEY (`IdLista_Proveedor`),
  KEY `fk_Lista_Proveedores_Lista_Productos1_idx` (`IdLista_Producto`),
  KEY `fk_Lista_Proveedores_Proveedores1_idx` (`Nit_Proveedor`),
  KEY `fk_Lista_Proveedores_Eval_Proveedores1_idx` (`IdEval_Proveedor`),
  CONSTRAINT `fk_Lista_Proveedores_Eval_Proveedores1` FOREIGN KEY (`IdEval_Proveedor`) REFERENCES `eval_proveedores` (`IdEval_Proveedor`) ON DELETE NO ACTION ON UPDATE NO ACTION,
  CONSTRAINT `fk_Lista_Proveedores_Lista_Productos1` FOREIGN KEY (`IdLista_Producto`) REFERENCES `lista_productos` (`IdLista_Productos`) ON DELETE NO ACTION ON UPDATE NO ACTION,
  CONSTRAINT `fk_Lista_Proveedores_Proveedores1` FOREIGN KEY (`Nit_Proveedor`) REFERENCES `proveedores` (`Nit_Proveedor`) ON DELETE NO ACTION ON UPDATE NO ACTION
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `lista_proveedores`
--

LOCK TABLES `lista_proveedores` WRITE;
/*!40000 ALTER TABLE `lista_proveedores` DISABLE KEYS */;
/*!40000 ALTER TABLE `lista_proveedores` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `movimientos`
--

DROP TABLE IF EXISTS `movimientos`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `movimientos` (
  `Id_Movimientos` int(11) NOT NULL AUTO_INCREMENT,
  `IdOrden_Movimiento` int(11) DEFAULT NULL,
  `Id_Producto` int(11) DEFAULT NULL,
  `Cantidad` int(11) DEFAULT NULL,
  `Precio_Compra` varchar(45) DEFAULT NULL,
  `Id_Usuario` int(11) DEFAULT NULL,
  `Fecha` date DEFAULT NULL,
  `Descripcion` varchar(145) DEFAULT NULL,
  `Id_Usuario_Final` int(11) DEFAULT NULL,
  PRIMARY KEY (`Id_Movimientos`),
  KEY `fk_Movimientos_Productos1_idx` (`Id_Producto`),
  KEY `fk_Movimientos_Usuarios1_idx` (`Id_Usuario`),
  KEY `fk_Movimientos_Orden_Movimientos1_idx` (`IdOrden_Movimiento`),
  CONSTRAINT `fk_Movimientos_Orden_Movimientos1` FOREIGN KEY (`IdOrden_Movimiento`) REFERENCES `orden_movimientos` (`IdOrden_Movimiento`) ON DELETE NO ACTION ON UPDATE NO ACTION,
  CONSTRAINT `fk_Movimientos_Productos1` FOREIGN KEY (`Id_Producto`) REFERENCES `productos` (`Id_Producto`) ON DELETE NO ACTION ON UPDATE NO ACTION,
  CONSTRAINT `fk_Movimientos_Usuarios1` FOREIGN KEY (`Id_Usuario`) REFERENCES `usuarios` (`Id_Usuario`) ON DELETE NO ACTION ON UPDATE NO ACTION
) ENGINE=InnoDB AUTO_INCREMENT=29 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `movimientos`
--

LOCK TABLES `movimientos` WRITE;
/*!40000 ALTER TABLE `movimientos` DISABLE KEYS */;
INSERT INTO `movimientos` VALUES (1,6,4,6,'6354',1,'2017-10-12','4TRH RJ RT',NULL),(2,9,4,8,'765343',1,'2017-10-12','56UJYRNGF',NULL),(3,12,4,3,NULL,1,'2017-10-12','TRGNV C ',1),(4,12,4,2,NULL,1,'2017-10-12','GRVEF',1),(5,13,5,7,'7856',1,'2017-10-12','RGFB G',NULL),(6,13,5,1,'546',1,'2017-10-12','JYR',NULL),(7,14,1,5,'65',1,'2017-10-12','ERT GFN ',NULL),(8,14,1,1,'95',1,'2017-10-12','THHBF',NULL),(9,16,5,5,'3543',1,'2017-10-12','6RTBY RT',NULL),(10,16,5,5,'765',1,'2017-10-12','TRNG',NULL),(11,17,5,7,NULL,1,'2017-10-12','RYT GN NTR',1),(12,17,5,1,NULL,1,'2017-10-12','RE',1),(13,18,4,8,'2346',1,'2017-10-12','N R',NULL),(14,19,4,3,'52',1,'2017-10-18','OULNK',NULL),(15,19,4,2,'433',1,'2017-10-18','THFDB',NULL),(16,20,5,7,'67',1,'2017-10-18','5BJ5',NULL),(17,21,5,4,NULL,1,'2017-10-18','3B6BJTY TYJ',5),(18,21,4,8,NULL,1,'2017-10-18','UH I ',4),(19,22,6,3,'1000000',1,'2017-10-21','ENGTRADA BASE',NULL),(20,22,5,150,'200',1,'2017-10-21','REPUESTOS EQUI[POS VARIOS',NULL),(21,22,1,10,'1500',1,'2017-10-21','PRUEBA INT',NULL),(22,23,5,50,NULL,1,'2017-10-21','USO MTTO',5),(23,23,6,1,NULL,1,'2017-10-21','SALIDA BALANZA',6),(24,23,1,4,NULL,1,'2017-10-21','PRUEBA',4),(25,24,1,1,NULL,1,'2017-10-21','INT',2),(26,24,5,20,NULL,1,'2017-10-21','THT',2),(27,25,5,10,NULL,1,'2017-10-21','DSTH',4),(28,26,5,10,NULL,1,'2017-10-21','ERGF',4);
/*!40000 ALTER TABLE `movimientos` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `orden_movimientos`
--

DROP TABLE IF EXISTS `orden_movimientos`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `orden_movimientos` (
  `IdOrden_Movimiento` int(11) NOT NULL AUTO_INCREMENT,
  `Tipo` varchar(45) DEFAULT NULL,
  `N_Orden_Compra` varchar(45) DEFAULT NULL,
  `N_Referencia` varchar(45) DEFAULT NULL,
  `Fecha` varchar(45) DEFAULT NULL,
  `Observaciones` varchar(45) DEFAULT NULL,
  `Monto` varchar(45) DEFAULT NULL,
  `Nit_Proveedor` varchar(45) DEFAULT NULL,
  `IdEval_Proveedor` int(11) DEFAULT NULL,
  PRIMARY KEY (`IdOrden_Movimiento`),
  KEY `fk_Orden_Movimientos_Proveedores1_idx` (`Nit_Proveedor`),
  KEY `fk_Orden_Movimientos_Eval_Proveedores1_idx` (`IdEval_Proveedor`),
  CONSTRAINT `fk_Orden_Movimientos_Eval_Proveedores1` FOREIGN KEY (`IdEval_Proveedor`) REFERENCES `eval_proveedores` (`IdEval_Proveedor`) ON DELETE NO ACTION ON UPDATE NO ACTION,
  CONSTRAINT `fk_Orden_Movimientos_Proveedores1` FOREIGN KEY (`Nit_Proveedor`) REFERENCES `proveedores` (`Nit_Proveedor`) ON DELETE NO ACTION ON UPDATE NO ACTION
) ENGINE=InnoDB AUTO_INCREMENT=27 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `orden_movimientos`
--

LOCK TABLES `orden_movimientos` WRITE;
/*!40000 ALTER TABLE `orden_movimientos` DISABLE KEYS */;
INSERT INTO `orden_movimientos` VALUES (1,'INGRESO','RYTJ456','HRT75','11-OCT-17','JGRTYGH HG NG','6578576','900100375-0',NULL),(2,'INGRESO','3653Y345','435G6Y34','11-OCT-17','TDBH NRTYGN','5643','900100375-0',NULL),(3,'INGRESO','436345UYTH','G4G65','12-OCT-17','H RTGFN RTNR NRN','4544563','1325432K',NULL),(4,'INGRESO','543TRE','543','12-OCT-17',' RTYGN FR ','76453','1325432K',NULL),(5,'INGRESO','254G24','46VY45','2017-10-12','TRHERT EYJ EY','7455','1325432K',NULL),(6,'INGRESO','345G6VY34','364YV','2017-10-12','RTEH T TY T','347564','1325432K',NULL),(7,'INGRESO','453YR3','346V324','2017-10-12',' GGFFG R R ','334352','1325432K',NULL),(8,'INGRESO','43RTVEG','346','2017-10-12','R JTYJY ','5435','1325432K',NULL),(9,'INGRESO','45RTH','4U6RTG','2017-10-12','NGF BVGF','867543','1325432K',NULL),(10,'SALIDA','R11141442709N071','R11141442709N071','2017-10-12','RTDF',NULL,NULL,NULL),(11,'SALIDA','R1551809040N071','R1551809040N071','2017-10-12','WGRED',NULL,NULL,NULL),(12,'SALIDA','R11695322304N071','R11695322304N071','2017-10-12','RWTH',NULL,NULL,NULL),(13,'INGRESO','45G7543','59J78','2017-10-12','R RTRYM T','65','900100375-0',NULL),(14,'INGRESO','7BU','7UB','2017-10-12',' YTB563 6','7564','900100375-0',NULL),(15,'SALIDA','R11350599591N071','R11350599591N071','2017-10-12','G',NULL,NULL,NULL),(16,'INGRESO','35TEF','76UTH','2017-10-12','FB  TFB','676453','900100375-0',NULL),(17,'SALIDA','R1494676638N071','R1494676638N071','2017-10-12','',NULL,NULL,NULL),(18,'INGRESO','543YV54','V5UV5','2017-10-12',' TYJETY','3247','1325432K',NULL),(19,'INGRESO','76UYTYH78OIU','O87IUH','2017-10-18','IUKB Y GU','7689','1325432K',NULL),(20,'INGRESO','456U54B','457G','2017-10-18','T TY JTY','5654','900100375-0',NULL),(21,'SALIDA','R11637943100N071','R11637943100N071','2017-10-18','35G4U5',NULL,NULL,NULL),(22,'INGRESO','INT001','REFINT001','2017-10-21','REGISTRO DE ENTRADA PRUEBA','4500000','900100375-0',NULL),(23,'SALIDA','R1263354569N071','R1263354569N071','2017-10-21','PRUEBA DE SALIDA',NULL,NULL,NULL),(24,'SALIDA','R11546550348N071','R11546550348N071','2017-10-21','PRU',NULL,NULL,NULL),(25,'SALIDA','S671781310','S671781310','2017-10-21','TR',NULL,NULL,NULL),(26,'SALIDA','S1733093445','S1733093445','2017-10-21','54',NULL,NULL,NULL);
/*!40000 ALTER TABLE `orden_movimientos` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `perfiles`
--

DROP TABLE IF EXISTS `perfiles`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `perfiles` (
  `Id_Perfil` int(11) NOT NULL AUTO_INCREMENT,
  `Nombre_Perfil` varchar(45) DEFAULT NULL,
  `Nivel_Permisos` int(11) DEFAULT NULL,
  PRIMARY KEY (`Id_Perfil`)
) ENGINE=InnoDB AUTO_INCREMENT=6 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `perfiles`
--

LOCK TABLES `perfiles` WRITE;
/*!40000 ALTER TABLE `perfiles` DISABLE KEYS */;
INSERT INTO `perfiles` VALUES (1,'Supervisor',0),(2,'Administrador',1),(3,'Almacenista',2),(4,'Analista',3),(5,'Comprador',4);
/*!40000 ALTER TABLE `perfiles` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `productos`
--

DROP TABLE IF EXISTS `productos`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `productos` (
  `Id_Producto` int(11) NOT NULL AUTO_INCREMENT,
  `Cod_Producto` varchar(45) DEFAULT NULL,
  `Nombre_Producto` varchar(55) DEFAULT NULL,
  `Id_Categoria` int(11) DEFAULT NULL,
  `Id_SubCategoria` int(11) DEFAULT NULL,
  `Marca` varchar(45) DEFAULT NULL,
  `Serie` varchar(45) DEFAULT NULL,
  `Foto` varchar(145) DEFAULT NULL,
  `Id_Stock` int(11) NOT NULL,
  `Activo` tinyint(1) DEFAULT NULL,
  PRIMARY KEY (`Id_Producto`,`Id_Stock`),
  KEY `fk_Productos_Categorias_Sub1_idx` (`Id_SubCategoria`),
  KEY `fk_Productos_Stock1_idx` (`Id_Stock`),
  KEY `fk_Productos_Categorias1_idx` (`Id_Categoria`),
  CONSTRAINT `fk_Productos_Categorias1` FOREIGN KEY (`Id_Categoria`) REFERENCES `categorias` (`Id_Categoria`) ON DELETE NO ACTION ON UPDATE NO ACTION,
  CONSTRAINT `fk_Productos_Categorias_Sub1` FOREIGN KEY (`Id_SubCategoria`) REFERENCES `categorias_sub` (`Id_SubCategoria`) ON DELETE NO ACTION ON UPDATE NO ACTION,
  CONSTRAINT `fk_Productos_Stock1` FOREIGN KEY (`Id_Stock`) REFERENCES `stock` (`Id_Stock`) ON DELETE NO ACTION ON UPDATE NO ACTION
) ENGINE=InnoDB AUTO_INCREMENT=7 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `productos`
--

LOCK TABLES `productos` WRITE;
/*!40000 ALTER TABLE `productos` DISABLE KEYS */;
INSERT INTO `productos` VALUES (1,'54YTGF563','PRUEBA P 1',3,6,'PRUEBA M','PRUEBA S','C:/Users/Msi/Desktop/Imagenes_Prueba/P_54YTGF563_1.png',1,0),(2,'345HGF3','PRUEBA P2',1,2,'MARC2','SERI2',NULL,2,1),(3,'FGVH485','SILLA',6,11,'RIMAX','7546HH',NULL,3,1),(4,'ETH85254','PRUEBA5',2,4,'GDHYTR','FDG5YR',NULL,4,0),(5,'145246548','FUSIBLE 20 AMP',7,18,'GATO','DFSSHD124','C:/Users/Msi/Desktop/Imagenes_Prueba/P_145246548_5.png',5,1),(6,'LP001','BALANZA ANALITICA',7,18,'TOLEDO','AJ123456','C:/Users/Msi/Desktop/Imagenes_Prueba/P_LP001_6.png',6,1);
/*!40000 ALTER TABLE `productos` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `proveedores`
--

DROP TABLE IF EXISTS `proveedores`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `proveedores` (
  `Nit_Proveedor` varchar(45) NOT NULL,
  `Nombre_Proveedor` varchar(45) DEFAULT NULL,
  `Nombre_Contacto` varchar(45) DEFAULT NULL,
  `Direccion` varchar(65) DEFAULT NULL,
  `Ciudad` varchar(45) DEFAULT NULL,
  `Numero_Telefono` varchar(45) DEFAULT NULL,
  `Email_Contacto` varchar(45) DEFAULT NULL,
  `Numero_Fax` varchar(45) DEFAULT NULL,
  `Pagina_Web` varchar(65) DEFAULT NULL,
  `Detalle` varchar(65) DEFAULT NULL,
  `Clasificacion_OIMS` varchar(45) DEFAULT NULL,
  `Aprovado` tinyint(1) DEFAULT NULL,
  `Adjunto` varchar(245) DEFAULT NULL,
  `Activo` tinyint(1) DEFAULT NULL,
  PRIMARY KEY (`Nit_Proveedor`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `proveedores`
--

LOCK TABLES `proveedores` WRITE;
/*!40000 ALTER TABLE `proveedores` DISABLE KEYS */;
INSERT INTO `proveedores` VALUES ('1325432K','PRUE 2','JAIME','CARRERA 4','CALI','4357346','PRU@EL.CO','3547467','WWW.PRU.COM','PRUEBA','SI',0,NULL,1),('900100375-0','IDETECO LTDA','HENRY SALINAS','CALLE 80','BOGOTA','4363402','IDETECOLTDA@IDETECO.COM','3176422019','WWW.IDETECO.COM','ASEGURAMIENTO METROLOGICO','SERVICIOS',1,NULL,1),('9001003759','METROIL','HENRY SALINAS','CRA 111C #88-05','BOGOTA','3176422019','HENRY.A.SALINAS@GMAIL.COM','3176422019','WWW.IDETECO.COM','SUMINISTROS ELECTRICOS','OM',1,NULL,1),('90897011','PRUEBA 1','PRUEBA CONTACTO','CALLE 8','BOGOTA DC','3653645','HIHORT@GER.CO','7689556','HEGRIGHEJNJ','EURHPGJRNGR','RIGUJG',1,'',1);
/*!40000 ALTER TABLE `proveedores` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `rel_productos_equipos`
--

DROP TABLE IF EXISTS `rel_productos_equipos`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `rel_productos_equipos` (
  `IdRel_Producto_Equipos` int(11) NOT NULL AUTO_INCREMENT,
  `Id_Producto` int(11) DEFAULT NULL,
  `Id_Equipo` int(11) DEFAULT NULL,
  PRIMARY KEY (`IdRel_Producto_Equipos`),
  KEY `fk_Rel_Productos_Equipos_Productos1_idx` (`Id_Producto`),
  KEY `fk_Rel_Productos_Equipos_Equipos1_idx` (`Id_Equipo`),
  CONSTRAINT `fk_Rel_Productos_Equipos_Equipos1` FOREIGN KEY (`Id_Equipo`) REFERENCES `equipos` (`Id_Equipo`) ON DELETE NO ACTION ON UPDATE NO ACTION,
  CONSTRAINT `fk_Rel_Productos_Equipos_Productos1` FOREIGN KEY (`Id_Producto`) REFERENCES `productos` (`Id_Producto`) ON DELETE NO ACTION ON UPDATE NO ACTION
) ENGINE=InnoDB AUTO_INCREMENT=10 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `rel_productos_equipos`
--

LOCK TABLES `rel_productos_equipos` WRITE;
/*!40000 ALTER TABLE `rel_productos_equipos` DISABLE KEYS */;
INSERT INTO `rel_productos_equipos` VALUES (1,1,2),(2,1,3),(3,3,1),(4,2,1),(7,5,6),(8,5,2),(9,6,6);
/*!40000 ALTER TABLE `rel_productos_equipos` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `rel_productos_proveedores`
--

DROP TABLE IF EXISTS `rel_productos_proveedores`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `rel_productos_proveedores` (
  `IdRel_Productos_Proveedores` int(11) NOT NULL AUTO_INCREMENT,
  `Id_Producto` int(11) DEFAULT NULL,
  `Nit_Proveedor` varchar(45) DEFAULT NULL,
  PRIMARY KEY (`IdRel_Productos_Proveedores`),
  KEY `fk_Rel_Productos_Proveedores_Productos1_idx` (`Id_Producto`),
  KEY `fk_Rel_Productos_Proveedores_Proveedores1_idx` (`Nit_Proveedor`),
  CONSTRAINT `fk_Rel_Productos_Proveedores_Productos1` FOREIGN KEY (`Id_Producto`) REFERENCES `productos` (`Id_Producto`) ON DELETE NO ACTION ON UPDATE NO ACTION,
  CONSTRAINT `fk_Rel_Productos_Proveedores_Proveedores1` FOREIGN KEY (`Nit_Proveedor`) REFERENCES `proveedores` (`Nit_Proveedor`) ON DELETE NO ACTION ON UPDATE NO ACTION
) ENGINE=InnoDB AUTO_INCREMENT=8 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `rel_productos_proveedores`
--

LOCK TABLES `rel_productos_proveedores` WRITE;
/*!40000 ALTER TABLE `rel_productos_proveedores` DISABLE KEYS */;
INSERT INTO `rel_productos_proveedores` VALUES (1,1,'900100375-0'),(2,4,'1325432K'),(4,5,'9001003759'),(5,5,'900100375-0'),(6,6,'9001003759'),(7,6,'900100375-0');
/*!40000 ALTER TABLE `rel_productos_proveedores` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `rel_ubicaciones_productos`
--

DROP TABLE IF EXISTS `rel_ubicaciones_productos`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `rel_ubicaciones_productos` (
  `IdRel_Ubicaciones_Productos` int(11) NOT NULL AUTO_INCREMENT,
  `Id_Ubicacion` int(11) DEFAULT NULL,
  `Id_Producto` int(11) DEFAULT NULL,
  `Cantidad` int(11) DEFAULT NULL,
  `Aforo` varchar(45) DEFAULT NULL,
  PRIMARY KEY (`IdRel_Ubicaciones_Productos`),
  KEY `fk_Rel_Ubicaciones_Productos_Ubicaciones1_idx` (`Id_Ubicacion`),
  KEY `fk_Rel_Ubicaciones_Productos_Productos1_idx` (`Id_Producto`),
  CONSTRAINT `fk_Rel_Ubicaciones_Productos_Productos1` FOREIGN KEY (`Id_Producto`) REFERENCES `productos` (`Id_Producto`) ON DELETE NO ACTION ON UPDATE NO ACTION,
  CONSTRAINT `fk_Rel_Ubicaciones_Productos_Ubicaciones1` FOREIGN KEY (`Id_Ubicacion`) REFERENCES `ubicaciones` (`Id_Ubicacion`) ON DELETE NO ACTION ON UPDATE NO ACTION
) ENGINE=InnoDB AUTO_INCREMENT=12 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `rel_ubicaciones_productos`
--

LOCK TABLES `rel_ubicaciones_productos` WRITE;
/*!40000 ALTER TABLE `rel_ubicaciones_productos` DISABLE KEYS */;
INSERT INTO `rel_ubicaciones_productos` VALUES (1,1,1,0,'300'),(2,2,1,3,'200'),(4,4,2,0,'N/A'),(7,2,5,75,'N/A'),(8,5,1,2,'N/A'),(9,4,4,3,'N/A'),(10,5,4,2,'N/A'),(11,2,6,2,'N/A');
/*!40000 ALTER TABLE `rel_ubicaciones_productos` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `solicitud_compras`
--

DROP TABLE IF EXISTS `solicitud_compras`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `solicitud_compras` (
  `IdSolicitud_Compra` int(11) NOT NULL AUTO_INCREMENT,
  `Id_Formato` varchar(45) DEFAULT NULL,
  `Id_Usuario` int(11) DEFAULT NULL,
  `Detalle` varchar(200) DEFAULT NULL,
  `Fecha` date DEFAULT NULL,
  `Formato_Moc` varchar(45) DEFAULT NULL,
  `Proveedor_Sugerido` varchar(45) DEFAULT NULL,
  `Proveedor_Seleccionado` varchar(45) DEFAULT NULL,
  `Justificacion` varchar(45) DEFAULT NULL,
  `Firma_Solicitud` varchar(45) DEFAULT NULL,
  `Firma_Aprovado` varchar(45) DEFAULT NULL,
  PRIMARY KEY (`IdSolicitud_Compra`),
  KEY `fk_Solicitud_Compras_Usuarios1_idx` (`Id_Usuario`),
  CONSTRAINT `fk_Solicitud_Compras_Usuarios1` FOREIGN KEY (`Id_Usuario`) REFERENCES `usuarios` (`Id_Usuario`) ON DELETE NO ACTION ON UPDATE NO ACTION
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `solicitud_compras`
--

LOCK TABLES `solicitud_compras` WRITE;
/*!40000 ALTER TABLE `solicitud_compras` DISABLE KEYS */;
/*!40000 ALTER TABLE `solicitud_compras` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `stock`
--

DROP TABLE IF EXISTS `stock`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `stock` (
  `Id_Stock` int(11) NOT NULL AUTO_INCREMENT,
  `Stock_Minimo` int(11) DEFAULT '0',
  `Stock_Maximo` int(11) DEFAULT '0',
  `Stock_Existente` int(11) DEFAULT '0',
  `Unidades` varchar(45) DEFAULT NULL,
  `Compra_Maxima` varchar(45) DEFAULT '0',
  PRIMARY KEY (`Id_Stock`)
) ENGINE=InnoDB AUTO_INCREMENT=7 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `stock`
--

LOCK TABLES `stock` WRITE;
/*!40000 ALTER TABLE `stock` DISABLE KEYS */;
INSERT INTO `stock` VALUES (1,200,650,5,'Kg','500'),(2,3,8,0,'Gal','7'),(3,0,100,0,'Uni','0'),(4,0,0,7,'Kg','0'),(5,100,1000,65,'Uni','0'),(6,10,150,2,'Uni','0');
/*!40000 ALTER TABLE `stock` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `ubicaciones`
--

DROP TABLE IF EXISTS `ubicaciones`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `ubicaciones` (
  `Id_Ubicacion` int(11) NOT NULL AUTO_INCREMENT,
  `Estante` varchar(45) DEFAULT NULL,
  `Entrepano` varchar(45) DEFAULT NULL,
  `Caja_Color` varchar(45) DEFAULT NULL,
  `Zona` varchar(45) DEFAULT NULL,
  PRIMARY KEY (`Id_Ubicacion`)
) ENGINE=InnoDB AUTO_INCREMENT=6 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `ubicaciones`
--

LOCK TABLES `ubicaciones` WRITE;
/*!40000 ALTER TABLE `ubicaciones` DISABLE KEYS */;
INSERT INTO `ubicaciones` VALUES (1,'3','2','Amarillo','Reactivos'),(2,'3','2','4','Bodega'),(3,'2','1','3','Gases'),(4,'1','1','Rojo','Acceso Especial'),(5,'2','4','Azul','Gases');
/*!40000 ALTER TABLE `ubicaciones` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `usuarios`
--

DROP TABLE IF EXISTS `usuarios`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `usuarios` (
  `Id_Usuario` int(11) NOT NULL AUTO_INCREMENT,
  `Nombre_Usuario` varchar(45) DEFAULT NULL,
  `Usuario` varchar(45) DEFAULT NULL,
  `Contrasena` varchar(45) DEFAULT NULL,
  `Id_Perfil` int(11) DEFAULT NULL,
  `Foto` text,
  `Id_Doag` int(11) DEFAULT NULL,
  PRIMARY KEY (`Id_Usuario`),
  KEY `fk_Usuarios_Doag1_idx` (`Id_Doag`,`Nombre_Usuario`),
  KEY `fk_Usuarios _Perfil1_idx` (`Id_Perfil`)
) ENGINE=InnoDB AUTO_INCREMENT=7 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `usuarios`
--

LOCK TABLES `usuarios` WRITE;
/*!40000 ALTER TABLE `usuarios` DISABLE KEYS */;
INSERT INTO `usuarios` VALUES (1,'Administrador','Admin','123',2,'C:/Users/Msi/Desktop/Imagenes_Prueba/U_Admin_1.png',3),(2,'Carlos Murcia','Carl.Mur','123',1,'C:/Users/Msi/Desktop/Imagenes_Prueba/U_Carl.Mur_2.png',3),(4,'Lucia Perez','Luci.97','432',4,'C:/Users/Msi/Desktop/Imagenes_Prueba/U_Luci.97_4.png',1),(5,'Camilo','Camilo90','123',3,NULL,1),(6,'MIguel Lopez','Lop.mig','123',5,'C:/Users/Msi/Desktop/Imagenes_Prueba/U_Lop.mig_6.png',4);
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

-- Dump completed on 2017-10-21 14:46:10
