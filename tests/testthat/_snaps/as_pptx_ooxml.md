# parse_to_ooxml(pptx) creates the correct nodes

    Code
      writeLines(as.character(xml))
    Output
      <a:p>
        <a:pPr>
          <a:spcBef>
            <a:spcPts val="0"/>
          </a:spcBef>
          <a:spcAft>
            <a:spcPts val="300"/>
          </a:spcAft>
        </a:pPr>
        <a:r>
          <a:rPr/>
          <a:t>hello</a:t>
        </a:r>
      </a:p>

# pptx ooxml can be generated from gt object

    Code
      writeLines(as.character(xml))
    Output
      <a:tbl>
        <a:tblPr firstRow="0" lastRow="0" firstCol="0" lastCol="0" bandCol="0" bandRow="0">
          <a:tableW type="auto" w="0"/>
        </a:tblPr>
        <a:tblGrid>
          <a:gridCol w="1016000"/>
          <a:gridCol w="1016000"/>
          <a:gridCol w="1016000"/>
          <a:gridCol w="1016000"/>
          <a:gridCol w="1016000"/>
          <a:gridCol w="1016000"/>
          <a:gridCol w="1016000"/>
          <a:gridCol w="1016000"/>
          <a:gridCol w="1016000"/>
        </a:tblGrid>
        <a:tr h="10000">
          <a:tc>
            <a:tcPr>
              <a:tcBdr>
                <a:lnT w="25400" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="25400" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnB>
                <a:lnL w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnL>
              </a:tcBdr>
            </a:tcPr>
            <a:txBody>
              <a:bodyPr/>
              <a:lstStyle/>
              <a:p>
                <a:pPr algn="r">
                  <a:spcBef>
                    <a:spcPts val="0"/>
                  </a:spcBef>
                  <a:spcAft>
                    <a:spcPts val="300"/>
                  </a:spcAft>
                </a:pPr>
                <a:r>
                  <a:rPr sz="1000">
                    <a:latin typeface="Calibri"/>
                  </a:rPr>
                  <a:t xml:space="default">num</a:t>
                </a:r>
              </a:p>
            </a:txBody>
          </a:tc>
          <a:tc>
            <a:tcPr>
              <a:tcBdr>
                <a:lnT w="25400" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="25400" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnB>
              </a:tcBdr>
            </a:tcPr>
            <a:txBody>
              <a:bodyPr/>
              <a:lstStyle/>
              <a:p>
                <a:pPr algn="l">
                  <a:spcBef>
                    <a:spcPts val="0"/>
                  </a:spcBef>
                  <a:spcAft>
                    <a:spcPts val="300"/>
                  </a:spcAft>
                </a:pPr>
                <a:r>
                  <a:rPr sz="1000">
                    <a:latin typeface="Calibri"/>
                  </a:rPr>
                  <a:t xml:space="default">char</a:t>
                </a:r>
              </a:p>
            </a:txBody>
          </a:tc>
          <a:tc>
            <a:tcPr>
              <a:tcBdr>
                <a:lnT w="25400" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="25400" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnB>
              </a:tcBdr>
            </a:tcPr>
            <a:txBody>
              <a:bodyPr/>
              <a:lstStyle/>
              <a:p>
                <a:pPr algn="ctr">
                  <a:spcBef>
                    <a:spcPts val="0"/>
                  </a:spcBef>
                  <a:spcAft>
                    <a:spcPts val="300"/>
                  </a:spcAft>
                </a:pPr>
                <a:r>
                  <a:rPr sz="1000">
                    <a:latin typeface="Calibri"/>
                  </a:rPr>
                  <a:t xml:space="default">fctr</a:t>
                </a:r>
              </a:p>
            </a:txBody>
          </a:tc>
          <a:tc>
            <a:tcPr>
              <a:tcBdr>
                <a:lnT w="25400" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="25400" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnB>
              </a:tcBdr>
            </a:tcPr>
            <a:txBody>
              <a:bodyPr/>
              <a:lstStyle/>
              <a:p>
                <a:pPr algn="r">
                  <a:spcBef>
                    <a:spcPts val="0"/>
                  </a:spcBef>
                  <a:spcAft>
                    <a:spcPts val="300"/>
                  </a:spcAft>
                </a:pPr>
                <a:r>
                  <a:rPr sz="1000">
                    <a:latin typeface="Calibri"/>
                  </a:rPr>
                  <a:t xml:space="default">date</a:t>
                </a:r>
              </a:p>
            </a:txBody>
          </a:tc>
          <a:tc>
            <a:tcPr>
              <a:tcBdr>
                <a:lnT w="25400" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="25400" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnB>
              </a:tcBdr>
            </a:tcPr>
            <a:txBody>
              <a:bodyPr/>
              <a:lstStyle/>
              <a:p>
                <a:pPr algn="r">
                  <a:spcBef>
                    <a:spcPts val="0"/>
                  </a:spcBef>
                  <a:spcAft>
                    <a:spcPts val="300"/>
                  </a:spcAft>
                </a:pPr>
                <a:r>
                  <a:rPr sz="1000">
                    <a:latin typeface="Calibri"/>
                  </a:rPr>
                  <a:t xml:space="default">time</a:t>
                </a:r>
              </a:p>
            </a:txBody>
          </a:tc>
          <a:tc>
            <a:tcPr>
              <a:tcBdr>
                <a:lnT w="25400" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="25400" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnB>
              </a:tcBdr>
            </a:tcPr>
            <a:txBody>
              <a:bodyPr/>
              <a:lstStyle/>
              <a:p>
                <a:pPr algn="r">
                  <a:spcBef>
                    <a:spcPts val="0"/>
                  </a:spcBef>
                  <a:spcAft>
                    <a:spcPts val="300"/>
                  </a:spcAft>
                </a:pPr>
                <a:r>
                  <a:rPr sz="1000">
                    <a:latin typeface="Calibri"/>
                  </a:rPr>
                  <a:t xml:space="default">datetime</a:t>
                </a:r>
              </a:p>
            </a:txBody>
          </a:tc>
          <a:tc>
            <a:tcPr>
              <a:tcBdr>
                <a:lnT w="25400" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="25400" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnB>
              </a:tcBdr>
            </a:tcPr>
            <a:txBody>
              <a:bodyPr/>
              <a:lstStyle/>
              <a:p>
                <a:pPr algn="r">
                  <a:spcBef>
                    <a:spcPts val="0"/>
                  </a:spcBef>
                  <a:spcAft>
                    <a:spcPts val="300"/>
                  </a:spcAft>
                </a:pPr>
                <a:r>
                  <a:rPr sz="1000">
                    <a:latin typeface="Calibri"/>
                  </a:rPr>
                  <a:t xml:space="default">currency</a:t>
                </a:r>
              </a:p>
            </a:txBody>
          </a:tc>
          <a:tc>
            <a:tcPr>
              <a:tcBdr>
                <a:lnT w="25400" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="25400" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnB>
              </a:tcBdr>
            </a:tcPr>
            <a:txBody>
              <a:bodyPr/>
              <a:lstStyle/>
              <a:p>
                <a:pPr algn="l">
                  <a:spcBef>
                    <a:spcPts val="0"/>
                  </a:spcBef>
                  <a:spcAft>
                    <a:spcPts val="300"/>
                  </a:spcAft>
                </a:pPr>
                <a:r>
                  <a:rPr sz="1000">
                    <a:latin typeface="Calibri"/>
                  </a:rPr>
                  <a:t xml:space="default">row</a:t>
                </a:r>
              </a:p>
            </a:txBody>
          </a:tc>
          <a:tc>
            <a:tcPr>
              <a:tcBdr>
                <a:lnT w="25400" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="25400" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnB>
                <a:lnR w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnR>
              </a:tcBdr>
            </a:tcPr>
            <a:txBody>
              <a:bodyPr/>
              <a:lstStyle/>
              <a:p>
                <a:pPr algn="l">
                  <a:spcBef>
                    <a:spcPts val="0"/>
                  </a:spcBef>
                  <a:spcAft>
                    <a:spcPts val="300"/>
                  </a:spcAft>
                </a:pPr>
                <a:r>
                  <a:rPr sz="1000">
                    <a:latin typeface="Calibri"/>
                  </a:rPr>
                  <a:t xml:space="default">group</a:t>
                </a:r>
              </a:p>
            </a:txBody>
          </a:tc>
        </a:tr>
        <a:tr h="10000">
          <a:tc>
            <a:tcPr>
              <a:tcBdr>
                <a:lnT w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnB>
                <a:lnL w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnL>
                <a:lnR w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnR>
              </a:tcBdr>
            </a:tcPr>
            <a:txBody>
              <a:bodyPr/>
              <a:lstStyle/>
              <a:p>
                <a:pPr algn="r">
                  <a:spcBef>
                    <a:spcPts val="0"/>
                  </a:spcBef>
                  <a:spcAft>
                    <a:spcPts val="300"/>
                  </a:spcAft>
                </a:pPr>
                <a:r>
                  <a:rPr sz="1000">
                    <a:latin typeface="Calibri"/>
                  </a:rPr>
                  <a:t xml:space="default">0.1111</a:t>
                </a:r>
              </a:p>
            </a:txBody>
          </a:tc>
          <a:tc>
            <a:tcPr>
              <a:tcBdr>
                <a:lnT w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnB>
                <a:lnL w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnL>
                <a:lnR w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnR>
              </a:tcBdr>
            </a:tcPr>
            <a:txBody>
              <a:bodyPr/>
              <a:lstStyle/>
              <a:p>
                <a:pPr algn="l">
                  <a:spcBef>
                    <a:spcPts val="0"/>
                  </a:spcBef>
                  <a:spcAft>
                    <a:spcPts val="300"/>
                  </a:spcAft>
                </a:pPr>
                <a:r>
                  <a:rPr sz="1000">
                    <a:latin typeface="Calibri"/>
                  </a:rPr>
                  <a:t xml:space="default">apricot</a:t>
                </a:r>
              </a:p>
            </a:txBody>
          </a:tc>
          <a:tc>
            <a:tcPr>
              <a:tcBdr>
                <a:lnT w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnB>
                <a:lnL w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnL>
                <a:lnR w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnR>
              </a:tcBdr>
            </a:tcPr>
            <a:txBody>
              <a:bodyPr/>
              <a:lstStyle/>
              <a:p>
                <a:pPr algn="ctr">
                  <a:spcBef>
                    <a:spcPts val="0"/>
                  </a:spcBef>
                  <a:spcAft>
                    <a:spcPts val="300"/>
                  </a:spcAft>
                </a:pPr>
                <a:r>
                  <a:rPr sz="1000">
                    <a:latin typeface="Calibri"/>
                  </a:rPr>
                  <a:t xml:space="default">one</a:t>
                </a:r>
              </a:p>
            </a:txBody>
          </a:tc>
          <a:tc>
            <a:tcPr>
              <a:tcBdr>
                <a:lnT w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnB>
                <a:lnL w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnL>
                <a:lnR w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnR>
              </a:tcBdr>
            </a:tcPr>
            <a:txBody>
              <a:bodyPr/>
              <a:lstStyle/>
              <a:p>
                <a:pPr algn="r">
                  <a:spcBef>
                    <a:spcPts val="0"/>
                  </a:spcBef>
                  <a:spcAft>
                    <a:spcPts val="300"/>
                  </a:spcAft>
                </a:pPr>
                <a:r>
                  <a:rPr sz="1000">
                    <a:latin typeface="Calibri"/>
                  </a:rPr>
                  <a:t xml:space="default">2015-01-15</a:t>
                </a:r>
              </a:p>
            </a:txBody>
          </a:tc>
          <a:tc>
            <a:tcPr>
              <a:tcBdr>
                <a:lnT w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnB>
                <a:lnL w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnL>
                <a:lnR w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnR>
              </a:tcBdr>
            </a:tcPr>
            <a:txBody>
              <a:bodyPr/>
              <a:lstStyle/>
              <a:p>
                <a:pPr algn="r">
                  <a:spcBef>
                    <a:spcPts val="0"/>
                  </a:spcBef>
                  <a:spcAft>
                    <a:spcPts val="300"/>
                  </a:spcAft>
                </a:pPr>
                <a:r>
                  <a:rPr sz="1000">
                    <a:latin typeface="Calibri"/>
                  </a:rPr>
                  <a:t xml:space="default">13:35</a:t>
                </a:r>
              </a:p>
            </a:txBody>
          </a:tc>
          <a:tc>
            <a:tcPr>
              <a:tcBdr>
                <a:lnT w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnB>
                <a:lnL w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnL>
                <a:lnR w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnR>
              </a:tcBdr>
            </a:tcPr>
            <a:txBody>
              <a:bodyPr/>
              <a:lstStyle/>
              <a:p>
                <a:pPr algn="r">
                  <a:spcBef>
                    <a:spcPts val="0"/>
                  </a:spcBef>
                  <a:spcAft>
                    <a:spcPts val="300"/>
                  </a:spcAft>
                </a:pPr>
                <a:r>
                  <a:rPr sz="1000">
                    <a:latin typeface="Calibri"/>
                  </a:rPr>
                  <a:t xml:space="default">2018-01-01 02:22</a:t>
                </a:r>
              </a:p>
            </a:txBody>
          </a:tc>
          <a:tc>
            <a:tcPr>
              <a:tcBdr>
                <a:lnT w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnB>
                <a:lnL w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnL>
                <a:lnR w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnR>
              </a:tcBdr>
            </a:tcPr>
            <a:txBody>
              <a:bodyPr/>
              <a:lstStyle/>
              <a:p>
                <a:pPr algn="r">
                  <a:spcBef>
                    <a:spcPts val="0"/>
                  </a:spcBef>
                  <a:spcAft>
                    <a:spcPts val="300"/>
                  </a:spcAft>
                </a:pPr>
                <a:r>
                  <a:rPr sz="1000">
                    <a:latin typeface="Calibri"/>
                  </a:rPr>
                  <a:t xml:space="default">49.95</a:t>
                </a:r>
              </a:p>
            </a:txBody>
          </a:tc>
          <a:tc>
            <a:tcPr>
              <a:tcBdr>
                <a:lnT w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnB>
                <a:lnL w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnL>
                <a:lnR w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnR>
              </a:tcBdr>
            </a:tcPr>
            <a:txBody>
              <a:bodyPr/>
              <a:lstStyle/>
              <a:p>
                <a:pPr algn="l">
                  <a:spcBef>
                    <a:spcPts val="0"/>
                  </a:spcBef>
                  <a:spcAft>
                    <a:spcPts val="300"/>
                  </a:spcAft>
                </a:pPr>
                <a:r>
                  <a:rPr sz="1000">
                    <a:latin typeface="Calibri"/>
                  </a:rPr>
                  <a:t xml:space="default">row_1</a:t>
                </a:r>
              </a:p>
            </a:txBody>
          </a:tc>
          <a:tc>
            <a:tcPr>
              <a:tcBdr>
                <a:lnT w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnB>
                <a:lnL w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnL>
                <a:lnR w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnR>
              </a:tcBdr>
            </a:tcPr>
            <a:txBody>
              <a:bodyPr/>
              <a:lstStyle/>
              <a:p>
                <a:pPr algn="l">
                  <a:spcBef>
                    <a:spcPts val="0"/>
                  </a:spcBef>
                  <a:spcAft>
                    <a:spcPts val="300"/>
                  </a:spcAft>
                </a:pPr>
                <a:r>
                  <a:rPr sz="1000">
                    <a:latin typeface="Calibri"/>
                  </a:rPr>
                  <a:t xml:space="default">grp_a</a:t>
                </a:r>
              </a:p>
            </a:txBody>
          </a:tc>
        </a:tr>
      </a:tbl>

---

    Code
      writeLines(as.character(xml))
    Output
      <a:p>
        <a:pPr algn="ctr">
          <a:spcBef>
            <a:spcPts val="0"/>
          </a:spcBef>
          <a:spcAft>
            <a:spcPts val="300"/>
          </a:spcAft>
        </a:pPr>
        <a:r>
          <a:rPr sz="1200">
            <a:latin typeface="Calibri"/>
            <a:solidFill>
              <a:srgbClr val="333333"/>
            </a:solidFill>
          </a:rPr>
          <a:t xml:space="default">TABLE TITLE</a:t>
        </a:r>
      </a:p>

---

    Code
      writeLines(as.character(xml))
    Output
      <a:p>
        <a:pPr algn="ctr">
          <a:spcBef>
            <a:spcPts val="0"/>
          </a:spcBef>
          <a:spcAft>
            <a:spcPts val="300"/>
          </a:spcAft>
        </a:pPr>
        <a:r>
          <a:rPr sz="800">
            <a:latin typeface="Calibri"/>
            <a:solidFill>
              <a:srgbClr val="333333"/>
            </a:solidFill>
          </a:rPr>
          <a:t xml:space="default">table subtitle</a:t>
        </a:r>
      </a:p>

# sub_small_vals() and sub_large_vals() are properly encoded

    Code
      writeLines(as.character(xml))
    Output
      <a:t xml:space="default">x</a:t>
      <a:t xml:space="default">&lt;0.01</a:t>
      <a:t xml:space="default">0.01</a:t>
      <a:t xml:space="default">≥100</a:t>

---

    Code
      writeLines(as.character(xml))
    Output
      <a:t xml:space="default">y</a:t>
      <a:t xml:space="default">&lt;</a:t>
      <a:t xml:space="default">%</a:t>
      <a:t xml:space="default">&gt;</a:t>

