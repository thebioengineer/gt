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
      <a:tbl xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <a:tblPr a:firstRow="0" a:lastRow="0" a:firstColumn="0" a:lastCol="0" a:bandCol="0" a:bandRow="0">
          <a:tableW a:type="auto"/>
        </a:tblPr>
        <a:tblGrid>
          <a:gridCol/>
          <a:gridCol/>
          <a:gridCol/>
          <a:gridCol/>
          <a:gridCol/>
          <a:gridCol/>
          <a:gridCol/>
          <a:gridCol/>
          <a:gridCol/>
        </a:tblGrid>
        <a:tr>
          <a:tc>
            <a:tcPr>
              <a:tcBdr>
                <a:lnT w="25400" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr>D3D3D3</a:srgbClr>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="25400" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr>D3D3D3</a:srgbClr>
                  </a:solidFill>
                </a:lnB>
                <a:lnL w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr>D3D3D3</a:srgbClr>
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
                    <a:srgbClr>D3D3D3</a:srgbClr>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="25400" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr>D3D3D3</a:srgbClr>
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
                    <a:srgbClr>D3D3D3</a:srgbClr>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="25400" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr>D3D3D3</a:srgbClr>
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
                    <a:srgbClr>D3D3D3</a:srgbClr>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="25400" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr>D3D3D3</a:srgbClr>
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
                    <a:srgbClr>D3D3D3</a:srgbClr>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="25400" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr>D3D3D3</a:srgbClr>
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
                    <a:srgbClr>D3D3D3</a:srgbClr>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="25400" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr>D3D3D3</a:srgbClr>
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
                    <a:srgbClr>D3D3D3</a:srgbClr>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="25400" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr>D3D3D3</a:srgbClr>
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
                    <a:srgbClr>D3D3D3</a:srgbClr>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="25400" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr>D3D3D3</a:srgbClr>
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
                    <a:srgbClr>D3D3D3</a:srgbClr>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="25400" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr>D3D3D3</a:srgbClr>
                  </a:solidFill>
                </a:lnB>
                <a:lnR w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr>D3D3D3</a:srgbClr>
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
        <a:tr>
          <a:tc>
            <a:tcPr>
              <a:tcBdr>
                <a:lnT w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr>D3D3D3</a:srgbClr>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr>D3D3D3</a:srgbClr>
                  </a:solidFill>
                </a:lnB>
                <a:lnL w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr>D3D3D3</a:srgbClr>
                  </a:solidFill>
                </a:lnL>
                <a:lnR w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr>D3D3D3</a:srgbClr>
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
                    <a:srgbClr>D3D3D3</a:srgbClr>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr>D3D3D3</a:srgbClr>
                  </a:solidFill>
                </a:lnB>
                <a:lnL w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr>D3D3D3</a:srgbClr>
                  </a:solidFill>
                </a:lnL>
                <a:lnR w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr>D3D3D3</a:srgbClr>
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
                    <a:srgbClr>D3D3D3</a:srgbClr>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr>D3D3D3</a:srgbClr>
                  </a:solidFill>
                </a:lnB>
                <a:lnL w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr>D3D3D3</a:srgbClr>
                  </a:solidFill>
                </a:lnL>
                <a:lnR w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr>D3D3D3</a:srgbClr>
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
                    <a:srgbClr>D3D3D3</a:srgbClr>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr>D3D3D3</a:srgbClr>
                  </a:solidFill>
                </a:lnB>
                <a:lnL w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr>D3D3D3</a:srgbClr>
                  </a:solidFill>
                </a:lnL>
                <a:lnR w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr>D3D3D3</a:srgbClr>
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
                    <a:srgbClr>D3D3D3</a:srgbClr>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr>D3D3D3</a:srgbClr>
                  </a:solidFill>
                </a:lnB>
                <a:lnL w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr>D3D3D3</a:srgbClr>
                  </a:solidFill>
                </a:lnL>
                <a:lnR w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr>D3D3D3</a:srgbClr>
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
                    <a:srgbClr>D3D3D3</a:srgbClr>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr>D3D3D3</a:srgbClr>
                  </a:solidFill>
                </a:lnB>
                <a:lnL w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr>D3D3D3</a:srgbClr>
                  </a:solidFill>
                </a:lnL>
                <a:lnR w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr>D3D3D3</a:srgbClr>
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
                    <a:srgbClr>D3D3D3</a:srgbClr>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr>D3D3D3</a:srgbClr>
                  </a:solidFill>
                </a:lnB>
                <a:lnL w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr>D3D3D3</a:srgbClr>
                  </a:solidFill>
                </a:lnL>
                <a:lnR w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr>D3D3D3</a:srgbClr>
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
                    <a:srgbClr>D3D3D3</a:srgbClr>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr>D3D3D3</a:srgbClr>
                  </a:solidFill>
                </a:lnB>
                <a:lnL w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr>D3D3D3</a:srgbClr>
                  </a:solidFill>
                </a:lnL>
                <a:lnR w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr>D3D3D3</a:srgbClr>
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
                    <a:srgbClr>D3D3D3</a:srgbClr>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr>D3D3D3</a:srgbClr>
                  </a:solidFill>
                </a:lnB>
                <a:lnL w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr>D3D3D3</a:srgbClr>
                  </a:solidFill>
                </a:lnL>
                <a:lnR w="6350" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr>D3D3D3</a:srgbClr>
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
            <a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
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
            <a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:srgbClr val="333333"/>
            </a:solidFill>
          </a:rPr>
          <a:t xml:space="default">table subtitle</a:t>
        </a:r>
      </a:p>

