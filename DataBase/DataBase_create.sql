-- MySQL dump 10.13  Distrib 8.0.16, for Win64 (x86_64)
--
-- Host: 127.0.0.1    Database: transportfinder
-- ------------------------------------------------------
-- Server version	8.0.16

/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
 SET NAMES utf8 ;
/*!40103 SET @OLD_TIME_ZONE=@@TIME_ZONE */;
/*!40103 SET TIME_ZONE='+00:00' */;
/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;
/*!40111 SET @OLD_SQL_NOTES=@@SQL_NOTES, SQL_NOTES=0 */;

--
-- Table structure for table `owners`
--

DROP TABLE IF EXISTS `owners`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
 SET character_set_client = utf8mb4 ;
CREATE TABLE `owners` (
  `Owner_id` int(10) unsigned NOT NULL,
  `INN` varchar(20) DEFAULT 'Н/Д',
  `Title` varchar(255) DEFAULT 'Н/Д',
  `Registred_at` date DEFAULT NULL,
  `License_number` varchar(20) DEFAULT 'Н/Д',
  `Reg_adress` varchar(255) DEFAULT 'Н/Д',
  `Implement_adress` varchar(255) DEFAULT 'Н/Д',
  `Risk_category` varchar(255) DEFAULT 'Н/Д',
  PRIMARY KEY (`Owner_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `owners-inspecs`
--

DROP TABLE IF EXISTS `owners-inspecs`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
 SET character_set_client = utf8mb4 ;
CREATE TABLE `owners-inspecs` (
  `id` int(11) NOT NULL,
  `owner_id` int(10) unsigned DEFAULT NULL,
  `inspec_id` int(10) unsigned DEFAULT NULL,
  PRIMARY KEY (`id`),
  KEY `owner_id_idx` (`owner_id`),
  KEY `inspec_id_idx` (`inspec_id`),
  CONSTRAINT `owners-inspecs_ibfk_1` FOREIGN KEY (`owner_id`) REFERENCES `owners` (`Owner_id`),
  CONSTRAINT `owners-inspecs_ibfk_2` FOREIGN KEY (`inspec_id`) REFERENCES `prosec_inspecs` (`Inspec_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `prosec_inspecs`
--

DROP TABLE IF EXISTS `prosec_inspecs`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
 SET character_set_client = utf8mb4 ;
CREATE TABLE `prosec_inspecs` (
  `Inspec_id` int(10) unsigned NOT NULL,
  `Starts_at` date DEFAULT NULL,
  `Duration_hours` int(11) DEFAULT NULL,
  `Purpose` varchar(255) DEFAULT 'Н/Д',
  `other_reason` varchar(255) DEFAULT 'Н/Д',
  `form_of_holding` varchar(20) DEFAULT 'Н/Д',
  `Performs_with` text,
  `Risk_category` varchar(255) DEFAULT 'Н/Д',
  PRIMARY KEY (`Inspec_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `transport`
--

DROP TABLE IF EXISTS `transport`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
 SET character_set_client = utf8mb4 ;
CREATE TABLE `transport` (
  `transport_id` int(10) unsigned NOT NULL,
  `VIN` varchar(20) DEFAULT 'Н/Д',
  `State_Registr_Mark` varchar(20) DEFAULT 'Н/Д',
  `Region` smallint(6) DEFAULT NULL,
  `Date_of_issue` date DEFAULT NULL,
  `pass_ser` varchar(20) DEFAULT 'Н/Д',
  `Ownership` varchar(20) DEFAULT 'Н/Д',
  `brand` varchar(100) DEFAULT 'Н/Д',
  PRIMARY KEY (`transport_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `transport-inspecs`
--

DROP TABLE IF EXISTS `transport-inspecs`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
 SET character_set_client = utf8mb4 ;
CREATE TABLE `transport-inspecs` (
  `id` int(10) unsigned NOT NULL,
  `transport_id` int(10) unsigned DEFAULT NULL,
  `inspec_id` int(10) unsigned DEFAULT NULL,
  PRIMARY KEY (`id`),
  KEY `inspec_id_idx` (`inspec_id`),
  KEY `transport_id_idx` (`transport_id`),
  CONSTRAINT `transport-inspecs_ibfk_1` FOREIGN KEY (`transport_id`) REFERENCES `transport` (`transport_id`),
  CONSTRAINT `transport-inspecs_ibfk_2` FOREIGN KEY (`inspec_id`) REFERENCES `prosec_inspecs` (`Inspec_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `transport-owners`
--

DROP TABLE IF EXISTS `transport-owners`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
 SET character_set_client = utf8mb4 ;
CREATE TABLE `transport-owners` (
  `id_T- O` int(10) unsigned NOT NULL,
  `owner_id` int(10) unsigned DEFAULT NULL,
  `transport_id` int(10) unsigned DEFAULT NULL,
  PRIMARY KEY (`id_T- O`),
  KEY `owner_id_idx` (`owner_id`),
  KEY `transport_id_idx` (`transport_id`),
  CONSTRAINT `owner_id` FOREIGN KEY (`owner_id`) REFERENCES `owners` (`Owner_id`),
  CONSTRAINT `transport_id` FOREIGN KEY (`transport_id`) REFERENCES `transport` (`transport_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;
/*!40103 SET TIME_ZONE=@OLD_TIME_ZONE */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;

-- Dump completed on 2020-06-25 17:17:24
