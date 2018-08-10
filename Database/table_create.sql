--CREATE TABLE `opt10081` (
--  `종목코드` varchar(6) NOT NULL,
--  `현재가` int(11) NOT NULL,
--  `거래량` int(11) NOT NULL,
--  `거래대금` int(11) NOT NULL,
--  `일자` datetime NOT NULL,
--  `시가` int(11) NOT NULL,
--  `고가` int(11) NOT NULL,
--  `저가` int(11) NOT NULL,
--  `수정주가구분` varchar(40) DEFAULT NULL,
--  `수정비율` varchar(40) DEFAULT NULL,
--  `대업종구분` varchar(40) DEFAULT NULL,
--  `소업종구분` varchar(40) DEFAULT NULL,
--  `종목정보` varchar(40) DEFAULT NULL,
--  `수정주가이벤트` varchar(40) DEFAULT NULL,
--  `전일종가` varchar(40) DEFAULT NULL,
--  PRIMARY KEY (`종목코드`,`일자`)
--) ENGINE=InnoDB DEFAULT CHARSET=utf8 ROW_FORMAT=COMPACT;
--
--CREATE TABLE `opt10079` (
--  `종목코드` varchar(40) NOT NULL,
--  `현재가` varchar(40) DEFAULT NULL,
--  `거래량` varchar(40) DEFAULT NULL,
--  `거래대금` varchar(40) DEFAULT NULL,
--  `일자` varchar(40) NOT NULL,
--  `시가` varchar(40) DEFAULT NULL,
--  `고가` varchar(40) DEFAULT NULL,
--  `저가` varchar(40) DEFAULT NULL,
--  `수정주가구분` varchar(40) DEFAULT NULL,
--  `수정비율` varchar(40) DEFAULT NULL,
--  `대업종구분` varchar(40) DEFAULT NULL,
--  `소업종구분` varchar(40) DEFAULT NULL,
--  `종목정보` varchar(40) DEFAULT NULL,
--  `수정주가이벤트` varchar(40) DEFAULT NULL,
--  `전일종가` varchar(40) DEFAULT NULL,
--  PRIMARY KEY (`종목코드`,`일자`)
--) ENGINE=InnoDB DEFAULT CHARSET=utf8 ROW_FORMAT=COMPACT;

CREATE TABLE `opt10081` (
  `symbol` varchar(6) NOT NULL,
  `cur_price` int(11) NOT NULL,
  `volume` int(11) NOT NULL,
  `volume_price` int(11) NOT NULL,
  `date` datetime NOT NULL,
  `open_price` int(11) NOT NULL,
  `high_price` int(11) NOT NULL,
  `low_price` int(11) NOT NULL,
  `modify_gubun` varchar(40) DEFAULT NULL,
  `modify_ratio` varchar(40) DEFAULT NULL,
  `big_gubun` varchar(40) DEFAULT NULL,
  `small_gubun` varchar(40) DEFAULT NULL,
  `stock_inform` varchar(40) DEFAULT NULL,
  `modify_event` varchar(40) DEFAULT NULL,
  `before_close` varchar(40) DEFAULT NULL,
  PRIMARY KEY (`code`,`date`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 ROW_FORMAT=COMPACT;

CREATE TABLE `opt10079` (
  `code` varchar(40) NOT NULL,
  `cur_price` varchar(40) DEFAULT NULL,
  `volume` varchar(40) DEFAULT NULL,
  `volume_price` varchar(40) DEFAULT NULL,
  `date` varchar(40) NOT NULL,
  `open_price` varchar(40) DEFAULT NULL,
  `high_price` varchar(40) DEFAULT NULL,
  `low_price` varchar(40) DEFAULT NULL,
  `modify_gubun` varchar(40) DEFAULT NULL,
  `modify_ratio` varchar(40) DEFAULT NULL,
  `big_gubun` varchar(40) DEFAULT NULL,
  `small_gubun` varchar(40) DEFAULT NULL,
  `stock_inform` varchar(40) DEFAULT NULL,
  `modify_event` varchar(40) DEFAULT NULL,
  `before_close` varchar(40) DEFAULT NULL,
  PRIMARY KEY (`code`,`date`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 ROW_FORMAT=COMPACT;
