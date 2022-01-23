Module ModCRC16


    ' Evalutate the 16-bit CRC (Cyclic Redundancy Checksum) of an array of bytes
    '
    ' If you omit the second argument, the entire array is considered

    Function CRC16(cp() As Byte, Optional ByVal Size As Long = -1) As UInt16
        Dim i As Long
        Dim fcs As Long
        Static fcstab(255) As Long

        Const pppinitfcs16 = &HFFFF& 'Initial FCS value

        If Size < 0 Then Size = UBound(cp) - LBound(cp) + 1

        If fcstab(1) = 0 Then
            ' Initialize array once and for all
            fcstab(0) = &H0&
            fcstab(1) = &H1189&
            fcstab(2) = &H2312&
            fcstab(3) = &H329B&
            fcstab(4) = &H4624&
            fcstab(5) = &H57AD&
            fcstab(6) = &H6536&
            fcstab(7) = &H74BF&
            fcstab(8) = &H8C48&
            fcstab(9) = &H9DC1&
            fcstab(10) = &HAF5A&
            fcstab(11) = &HBED3&
            fcstab(12) = &HCA6C&
            fcstab(13) = &HDBE5&
            fcstab(14) = &HE97E&
            fcstab(15) = &HF8F7&
            fcstab(16) = &H1081&
            fcstab(17) = &H108&
            fcstab(18) = &H3393&
            fcstab(19) = &H221A&
            fcstab(20) = &H56A5&
            fcstab(21) = &H472C&
            fcstab(22) = &H75B7&
            fcstab(23) = &H643E&
            fcstab(24) = &H9CC9&
            fcstab(25) = &H8D40&
            fcstab(26) = &HBFDB&
            fcstab(27) = &HAE52&
            fcstab(28) = &HDAED&
            fcstab(29) = &HCB64&
            fcstab(30) = &HF9FF&
            fcstab(31) = &HE876&
            fcstab(32) = &H2102&
            fcstab(33) = &H308B&
            fcstab(34) = &H210&
            fcstab(35) = &H1399&
            fcstab(36) = &H6726&
            fcstab(37) = &H76AF&
            fcstab(38) = &H4434&
            fcstab(39) = &H55BD&
            fcstab(40) = &HAD4A&
            fcstab(41) = &HBCC3&
            fcstab(42) = &H8E58&
            fcstab(43) = &H9FD1&
            fcstab(44) = &HEB6E&
            fcstab(45) = &HFAE7&
            fcstab(46) = &HC87C&
            fcstab(47) = &HD9F5&
            fcstab(48) = &H3183&
            fcstab(49) = &H200A&
            fcstab(50) = &H1291&
            fcstab(51) = &H318&
            fcstab(52) = &H77A7&
            fcstab(53) = &H662E&
            fcstab(54) = &H54B5&
            fcstab(55) = &H453C&
            fcstab(56) = &HBDCB&
            fcstab(57) = &HAC42&
            fcstab(58) = &H9ED9&
            fcstab(59) = &H8F50&
            fcstab(60) = &HFBEF&
            fcstab(61) = &HEA66&
            fcstab(62) = &HD8FD&
            fcstab(63) = &HC974&
            fcstab(64) = &H4204&
            fcstab(65) = &H538D&
            fcstab(66) = &H6116&
            fcstab(67) = &H709F&
            fcstab(68) = &H420&
            fcstab(69) = &H15A9&
            fcstab(70) = &H2732&
            fcstab(71) = &H36BB&
            fcstab(72) = &HCE4C&
            fcstab(73) = &HDFC5&
            fcstab(74) = &HED5E&
            fcstab(75) = &HFCD7&
            fcstab(76) = &H8868&
            fcstab(77) = &H99E1&
            fcstab(78) = &HAB7A&
            fcstab(79) = &HBAF3&
            fcstab(80) = &H5285&
            fcstab(81) = &H430C&
            fcstab(82) = &H7197&
            fcstab(83) = &H601E&
            fcstab(84) = &H14A1&
            fcstab(85) = &H528&
            fcstab(86) = &H37B3&
            fcstab(87) = &H263A&
            fcstab(88) = &HDECD&
            fcstab(89) = &HCF44&
            fcstab(90) = &HFDDF&
            fcstab(91) = &HEC56&
            fcstab(92) = &H98E9&
            fcstab(93) = &H8960&
            fcstab(94) = &HBBFB&
            fcstab(95) = &HAA72&
            fcstab(96) = &H6306&
            fcstab(97) = &H728F&
            fcstab(98) = &H4014&
            fcstab(99) = &H519D&
            fcstab(100) = &H2522&
            fcstab(101) = &H34AB&
            fcstab(102) = &H630&
            fcstab(103) = &H17B9&
            fcstab(104) = &HEF4E&
            fcstab(105) = &HFEC7&
            fcstab(106) = &HCC5C&
            fcstab(107) = &HDDD5&
            fcstab(108) = &HA96A&
            fcstab(109) = &HB8E3&
            fcstab(110) = &H8A78&
            fcstab(111) = &H9BF1&
            fcstab(112) = &H7387&
            fcstab(113) = &H620E&
            fcstab(114) = &H5095&
            fcstab(115) = &H411C&
            fcstab(116) = &H35A3&
            fcstab(117) = &H242A&
            fcstab(118) = &H16B1&
            fcstab(119) = &H738&
            fcstab(120) = &HFFCF&
            fcstab(121) = &HEE46&
            fcstab(122) = &HDCDD&
            fcstab(123) = &HCD54&
            fcstab(124) = &HB9EB&
            fcstab(125) = &HA862&
            fcstab(126) = &H9AF9&
            fcstab(127) = &H8B70&
            fcstab(128) = &H8408&
            fcstab(129) = &H9581&
            fcstab(130) = &HA71A&
            fcstab(131) = &HB693&
            fcstab(132) = &HC22C&
            fcstab(133) = &HD3A5&
            fcstab(134) = &HE13E&
            fcstab(135) = &HF0B7&
            fcstab(136) = &H840&
            fcstab(137) = &H19C9&
            fcstab(138) = &H2B52&
            fcstab(139) = &H3ADB&
            fcstab(140) = &H4E64&
            fcstab(141) = &H5FED&
            fcstab(142) = &H6D76&
            fcstab(143) = &H7CFF&
            fcstab(144) = &H9489&
            fcstab(145) = &H8500&
            fcstab(146) = &HB79B&
            fcstab(147) = &HA612&
            fcstab(148) = &HD2AD&
            fcstab(149) = &HC324&
            fcstab(150) = &HF1BF&
            fcstab(151) = &HE036&
            fcstab(152) = &H18C1&
            fcstab(153) = &H948&
            fcstab(154) = &H3BD3&
            fcstab(155) = &H2A5A&
            fcstab(156) = &H5EE5&
            fcstab(157) = &H4F6C&
            fcstab(158) = &H7DF7&
            fcstab(159) = &H6C7E&
            fcstab(160) = &HA50A&
            fcstab(161) = &HB483&
            fcstab(162) = &H8618&
            fcstab(163) = &H9791&
            fcstab(164) = &HE32E&
            fcstab(165) = &HF2A7&
            fcstab(166) = &HC03C&
            fcstab(167) = &HD1B5&
            fcstab(168) = &H2942&
            fcstab(169) = &H38CB&
            fcstab(170) = &HA50&
            fcstab(171) = &H1BD9&
            fcstab(172) = &H6F66&
            fcstab(173) = &H7EEF&
            fcstab(174) = &H4C74&
            fcstab(175) = &H5DFD&
            fcstab(176) = &HB58B&
            fcstab(177) = &HA402&
            fcstab(178) = &H9699&
            fcstab(179) = &H8710&
            fcstab(180) = &HF3AF&
            fcstab(181) = &HE226&
            fcstab(182) = &HD0BD&
            fcstab(183) = &HC134&
            fcstab(184) = &H39C3&
            fcstab(185) = &H284A&
            fcstab(186) = &H1AD1&
            fcstab(187) = &HB58&
            fcstab(188) = &H7FE7&
            fcstab(189) = &H6E6E&
            fcstab(190) = &H5CF5&
            fcstab(191) = &H4D7C&
            fcstab(192) = &HC60C&
            fcstab(193) = &HD785&
            fcstab(194) = &HE51E&
            fcstab(195) = &HF497&
            fcstab(196) = &H8028&
            fcstab(197) = &H91A1&
            fcstab(198) = &HA33A&
            fcstab(199) = &HB2B3&
            fcstab(200) = &H4A44&
            fcstab(201) = &H5BCD&
            fcstab(202) = &H6956&
            fcstab(203) = &H78DF&
            fcstab(204) = &HC60&
            fcstab(205) = &H1DE9&
            fcstab(206) = &H2F72&
            fcstab(207) = &H3EFB&
            fcstab(208) = &HD68D&
            fcstab(209) = &HC704&
            fcstab(210) = &HF59F&
            fcstab(211) = &HE416&
            fcstab(212) = &H90A9&
            fcstab(213) = &H8120&
            fcstab(214) = &HB3BB&
            fcstab(215) = &HA232&
            fcstab(216) = &H5AC5&
            fcstab(217) = &H4B4C&
            fcstab(218) = &H79D7&
            fcstab(219) = &H685E&
            fcstab(220) = &H1CE1&
            fcstab(221) = &HD68&
            fcstab(222) = &H3FF3&
            fcstab(223) = &H2E7A&
            fcstab(224) = &HE70E&
            fcstab(225) = &HF687&
            fcstab(226) = &HC41C&
            fcstab(227) = &HD595&
            fcstab(228) = &HA12A&
            fcstab(229) = &HB0A3&
            fcstab(230) = &H8238&
            fcstab(231) = &H93B1&
            fcstab(232) = &H6B46&
            fcstab(233) = &H7ACF&
            fcstab(234) = &H4854&
            fcstab(235) = &H59DD&
            fcstab(236) = &H2D62&
            fcstab(237) = &H3CEB&
            fcstab(238) = &HE70&
            fcstab(239) = &H1FF9&
            fcstab(240) = &HF78F&
            fcstab(241) = &HE606&
            fcstab(242) = &HD49D&
            fcstab(243) = &HC514&
            fcstab(244) = &HB1AB&
            fcstab(245) = &HA022&
            fcstab(246) = &H92B9&
            fcstab(247) = &H8330&
            fcstab(248) = &H7BC7&
            fcstab(249) = &H6A4E&
            fcstab(250) = &H58D5&
            fcstab(251) = &H495C&
            fcstab(252) = &H3DE3&
            fcstab(253) = &H2C6A&
            fcstab(254) = &H1EF1&
            fcstab(255) = &HF78&
        End If

        ' The initial FCS value
        fcs = pppinitfcs16

        ' evaluate the FCS
        For i = LBound(cp) To LBound(cp) + Size - 1
            fcs = (fcs \ &H100&) Xor fcstab((fcs Xor cp(i)) And &HFF&)
        Next i

        ' return the result
        Return fcs And &HFFFF
    End Function



    Public Function fcs16(data As Byte()) As Long

        Dim CRCArray As UInt32()
        Dim i As Long
        Dim fcs As Long
        Dim fcsh As Long
        Dim fcsl As Long
        Dim ArrayPointer As Integer
        Dim ArrayElement As Long
        Dim fcscompl As Long

        CRCArray = {&H0, &H1189, &H2312, &H329B, &H4624, &H57AD, &H6536, &H74BF, &H8C48,
        &H9DC1, &HAF5A, &HBED3, &HCA6C, &HDBE5, &HE97E, &HF8F7,
             &H1081, &H108, &H3393, &H221A, &H56A5, &H472C, &H75B7, &H643E, &H9CC9,
        &H8D40, &HBFDB, &HAE52, &HDAED, &HCB64, &HF9FF, &HE876,
             &H2102, &H308B, &H210, &H1399, &H6726, &H76AF, &H4434, &H55BD, &HAD4A,
        &HBCC3, &H8E58, &H9FD1, &HEB6E, &HFAE7, &HC87C, &HD9F5,
             &H3183, &H200A, &H1291, &H318, &H77A7, &H662E, &H54B5, &H453C, &HBDCB,
        &HAC42, &H9ED9, &H8F50, &HFBEF, &HEA66, &HD8FD, &HC974,
             &H4204, &H538D, &H6116, &H709F, &H420, &H15A9, &H2732, &H36BB, &HCE4C,
        &HDFC5, &HED5E, &HFCD7, &H8868, &H99E1, &HAB7A, &HBAF3,
             &H5285, &H430C, &H7197, &H601E, &H14A1, &H528, &H37B3, &H263A, &HDECD,
        &HCF44, &HFDDF, &HEC56, &H98E9, &H8960, &HBBFB, &HAA72,
             &H6306, &H728F, &H4014, &H519D, &H2522, &H34AB, &H630, &H17B9, &HEF4E,
        &HFEC7, &HCC5C, &HDDD5, &HA96A, &HB8E3, &H8A78, &H9BF1,
             &H7387, &H620E, &H5095, &H411C, &H35A3, &H242A, &H16B1, &H738, &HFFCF,
        &HEE46, &HDCDD, &HCD54, &HB9EB, &HA862, &H9AF9, &H8B70,
             &H8408, &H9581, &HA71A, &HB693, &HC22C, &HD3A5, &HE13E, &HF0B7, &H840,
        &H19C9, &H2B52, &H3ADB, &H4E64, &H5FED, &H6D76, &H7CFF,
             &H9489, &H8500, &HB79B, &HA612, &HD2AD, &HC324, &HF1BF, &HE036, &H18C1, &H948, &H3BD3, &H2A5A, &H5EE5, &H4F6C, &H7DF7, &H6C7E,
             &HA50A, &HB483, &H8618, &H9791, &HE32E, &HF2A7, &HC03C, &HD1B5, &H2942, &H38CB, &HA50, &H1BD9, &H6F66, &H7EEF, &H4C74, &H5DFD,
             &HB58B, &HA402, &H9699, &H8710, &HF3AF, &HE226, &HD0BD, &HC134, &H39C3, &H284A, &H1AD1, &HB58, &H7FE7, &H6E6E, &H5CF5, &H4D7C,
             &HC60C, &HD785, &HE51E, &HF497, &H8028, &H91A1, &HA33A, &HB2B3, &H4A44, &H5BCD, &H6956, &H78DF, &HC60, &H1DE9, &H2F72, &H3EFB, &HD68D, &HC704, &HF59F, &HE416, &H90A9, &H8120, &HB3BB, &HA232, &H5AC5, &H4B4C, &H79D7, &H685E, &H1CE1, &HD68, &H3FF3, &H2E7A,
             &HE70E, &HF687, &HC41C, &HD595, &HA12A, &HB0A3, &H8238, &H93B1, &H6B46, &H7ACF, &H4854, &H59DD, &H2D62, &H3CEB, &HE70, &H1FF9,
             &HF78F, &HE606, &HD49D, &HC514, &HB1AB, &HA022, &H92B9, &H8330, &H7BC7, &H6A4E, &H58D5, &H495C, &H3DE3, &H2C6A, &H1EF1, &HF78}

        fcs = 65535

        For i = 0 To data.Length - 1

            'get value from Array
            ArrayPointer = (fcs Xor data(i)) And &HFF
            ArrayElement = CRCArray(ArrayPointer)
            'If ArrayElement < 0 Then  'values >&H8000 will be returned as a negativevalue(bug in VB?)
            'ArrayElement = 32768 - Math.Abs(ArrayElement) + 32768 'calculate validvalue from the negative value
            'End If

            'shifting fcs
            fcs = Int(fcs / 256)

            'calculate fcs
            fcs = fcs Xor ArrayElement

        Next i

        'build complement
        fcscompl = fcs Xor 65535

        'make fcs high byte and low byte
        fcsh = fcscompl And &HFF
        fcscompl = Int(fcscompl / 256)
        fcsl = fcscompl And &HFF

        'build fcs LSB-MSB
        fcs = fcsh * 256 + fcsl

        Return fcs

    End Function

End Module
