

 Server: localhost -  Database: eia_master_set_control -  Table: personal_login 

-- phpMyAdmin SQL Dump
-- version 2.10.3
-- http://www.phpmyadmin.net
-- 
-- Host: localhost
-- Generation Time: Mar 26, 2019 at 02:34 PM
-- Server version: 5.0.51
-- PHP Version: 5.2.6

SET SQL_MODE="NO_AUTO_VALUE_ON_ZERO";

-- 
-- Database: `eia_master_set_control`
-- 

-- --------------------------------------------------------

-- 
-- Table structure for table `personal_login`
-- 

CREATE TABLE `personal_login` (
  `userID` varchar(11) character set utf8 collate utf8_unicode_ci NOT NULL,
  `name` varchar(250) character set utf8 collate utf8_unicode_ci NOT NULL,
  `Group` varchar(250) character set utf8 collate utf8_unicode_ci NOT NULL
) ENGINE=MyISAM DEFAULT CHARSET=utf8;

-- 
-- Dumping data for table `personal_login`
-- 

INSERT INTO `personal_login` VALUES ('5020001830', 'Mr.Songglod Punvichartkul', 'SSC');
INSERT INTO `personal_login` VALUES ('5022123000', 'Miss.Natchanicha Sripool', 'SSC');
INSERT INTO `personal_login` VALUES ('5022160200', 'Miss.Thapanee Lakkanathin', 'SSC');
INSERT INTO `personal_login` VALUES ('5022808592', 'Mr.Kiattisak Phanphu', 'ETC');
INSERT INTO `personal_login` VALUES ('5022004162', 'Mr.Sudrak  Namee', 'ETC');
INSERT INTO `personal_login` VALUES ('5022181082', 'Mr.Nontacha Waisayarungruang', 'SSC');
INSERT INTO `personal_login` VALUES ('5020001830', 'Mr.Songglod Punvichartkul', 'ETC');
INSERT INTO `personal_login` VALUES ('5022123000', 'Miss.Natchanicha Sripool', 'ETC');
INSERT INTO `personal_login` VALUES ('5022160200', 'Miss.Thapanee Lakkanathin', 'ETC');
INSERT INTO `personal_login` VALUES ('5022060295', 'Miss. Anusara Srisuk', 'EIC');
  
  