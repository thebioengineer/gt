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
          <a:rPr lang="en-US"/>
          <a:t>hello</a:t>
        </a:r>
      </a:p>

# pptx ooxml can be generated from gt object

    Code
      writeLines(as.character(xml))
    Output
      <a:tbl>
        <a:tblPr firstRow="0" lastRow="0" firstCol="0" lastCol="0" bandCol="0" bandRow="0">
          <a:tableW type="auto" w="9144000"/>
          <a:tblLook firstRow="0" lastRow="0" firstCol="0" lastCol="0" noHBand="1" noVBand="1" val="04A0"/>
          <a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId>
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
        <a:tr h="254000">
          <a:tc>
            <a:txBody>
              <a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" rtlCol="0" anchor="ctr"/>
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
                  <a:rPr lang="en-US" sz="1000">
                    <a:latin typeface="Calibri"/>
                  </a:rPr>
                  <a:t xml:space="default">num</a:t>
                </a:r>
                <a:endParaRPr sz="1000" lang="en-US"/>
              </a:p>
            </a:txBody>
            <a:tcPr marT="45720" marB="45720" marL="45720" marR="45720" anchor="ctr">
              <a:tcBdr>
                <a:lnT w="19050" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="19050" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnB>
                <a:lnL w="4762" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnL>
              </a:tcBdr>
              <a:noFill/>
            </a:tcPr>
          </a:tc>
          <a:tc>
            <a:txBody>
              <a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" rtlCol="0" anchor="ctr"/>
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
                  <a:rPr lang="en-US" sz="1000">
                    <a:latin typeface="Calibri"/>
                  </a:rPr>
                  <a:t xml:space="default">char</a:t>
                </a:r>
                <a:endParaRPr sz="1000" lang="en-US"/>
              </a:p>
            </a:txBody>
            <a:tcPr marT="45720" marB="45720" marL="45720" marR="45720" anchor="ctr">
              <a:tcBdr>
                <a:lnT w="19050" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="19050" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnB>
              </a:tcBdr>
              <a:noFill/>
            </a:tcPr>
          </a:tc>
          <a:tc>
            <a:txBody>
              <a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" rtlCol="0" anchor="ctr"/>
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
                  <a:rPr lang="en-US" sz="1000">
                    <a:latin typeface="Calibri"/>
                  </a:rPr>
                  <a:t xml:space="default">fctr</a:t>
                </a:r>
                <a:endParaRPr sz="1000" lang="en-US"/>
              </a:p>
            </a:txBody>
            <a:tcPr marT="45720" marB="45720" marL="45720" marR="45720" anchor="ctr">
              <a:tcBdr>
                <a:lnT w="19050" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="19050" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnB>
              </a:tcBdr>
              <a:noFill/>
            </a:tcPr>
          </a:tc>
          <a:tc>
            <a:txBody>
              <a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" rtlCol="0" anchor="ctr"/>
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
                  <a:rPr lang="en-US" sz="1000">
                    <a:latin typeface="Calibri"/>
                  </a:rPr>
                  <a:t xml:space="default">date</a:t>
                </a:r>
                <a:endParaRPr sz="1000" lang="en-US"/>
              </a:p>
            </a:txBody>
            <a:tcPr marT="45720" marB="45720" marL="45720" marR="45720" anchor="ctr">
              <a:tcBdr>
                <a:lnT w="19050" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="19050" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnB>
              </a:tcBdr>
              <a:noFill/>
            </a:tcPr>
          </a:tc>
          <a:tc>
            <a:txBody>
              <a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" rtlCol="0" anchor="ctr"/>
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
                  <a:rPr lang="en-US" sz="1000">
                    <a:latin typeface="Calibri"/>
                  </a:rPr>
                  <a:t xml:space="default">time</a:t>
                </a:r>
                <a:endParaRPr sz="1000" lang="en-US"/>
              </a:p>
            </a:txBody>
            <a:tcPr marT="45720" marB="45720" marL="45720" marR="45720" anchor="ctr">
              <a:tcBdr>
                <a:lnT w="19050" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="19050" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnB>
              </a:tcBdr>
              <a:noFill/>
            </a:tcPr>
          </a:tc>
          <a:tc>
            <a:txBody>
              <a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" rtlCol="0" anchor="ctr"/>
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
                  <a:rPr lang="en-US" sz="1000">
                    <a:latin typeface="Calibri"/>
                  </a:rPr>
                  <a:t xml:space="default">datetime</a:t>
                </a:r>
                <a:endParaRPr sz="1000" lang="en-US"/>
              </a:p>
            </a:txBody>
            <a:tcPr marT="45720" marB="45720" marL="45720" marR="45720" anchor="ctr">
              <a:tcBdr>
                <a:lnT w="19050" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="19050" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnB>
              </a:tcBdr>
              <a:noFill/>
            </a:tcPr>
          </a:tc>
          <a:tc>
            <a:txBody>
              <a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" rtlCol="0" anchor="ctr"/>
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
                  <a:rPr lang="en-US" sz="1000">
                    <a:latin typeface="Calibri"/>
                  </a:rPr>
                  <a:t xml:space="default">currency</a:t>
                </a:r>
                <a:endParaRPr sz="1000" lang="en-US"/>
              </a:p>
            </a:txBody>
            <a:tcPr marT="45720" marB="45720" marL="45720" marR="45720" anchor="ctr">
              <a:tcBdr>
                <a:lnT w="19050" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="19050" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnB>
              </a:tcBdr>
              <a:noFill/>
            </a:tcPr>
          </a:tc>
          <a:tc>
            <a:txBody>
              <a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" rtlCol="0" anchor="ctr"/>
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
                  <a:rPr lang="en-US" sz="1000">
                    <a:latin typeface="Calibri"/>
                  </a:rPr>
                  <a:t xml:space="default">row</a:t>
                </a:r>
                <a:endParaRPr sz="1000" lang="en-US"/>
              </a:p>
            </a:txBody>
            <a:tcPr marT="45720" marB="45720" marL="45720" marR="45720" anchor="ctr">
              <a:tcBdr>
                <a:lnT w="19050" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="19050" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnB>
              </a:tcBdr>
              <a:noFill/>
            </a:tcPr>
          </a:tc>
          <a:tc>
            <a:txBody>
              <a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" rtlCol="0" anchor="ctr"/>
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
                  <a:rPr lang="en-US" sz="1000">
                    <a:latin typeface="Calibri"/>
                  </a:rPr>
                  <a:t xml:space="default">group</a:t>
                </a:r>
                <a:endParaRPr sz="1000" lang="en-US"/>
              </a:p>
            </a:txBody>
            <a:tcPr marT="45720" marB="45720" marL="45720" marR="45720" anchor="ctr">
              <a:tcBdr>
                <a:lnT w="19050" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="19050" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnB>
                <a:lnR w="4762" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnR>
              </a:tcBdr>
              <a:noFill/>
            </a:tcPr>
          </a:tc>
        </a:tr>
        <a:tr h="254000">
          <a:tc>
            <a:txBody>
              <a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" rtlCol="0" anchor="ctr"/>
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
                  <a:rPr lang="en-US" sz="1000">
                    <a:latin typeface="Calibri"/>
                  </a:rPr>
                  <a:t xml:space="default">0.1111</a:t>
                </a:r>
                <a:endParaRPr sz="1000" lang="en-US"/>
              </a:p>
            </a:txBody>
            <a:tcPr marT="45720" marB="45720" marL="45720" marR="45720" anchor="ctr">
              <a:tcBdr>
                <a:lnT w="4762" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="4762" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnB>
                <a:lnL w="4762" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnL>
                <a:lnR w="4762" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnR>
              </a:tcBdr>
              <a:noFill/>
            </a:tcPr>
          </a:tc>
          <a:tc>
            <a:txBody>
              <a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" rtlCol="0" anchor="ctr"/>
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
                  <a:rPr lang="en-US" sz="1000">
                    <a:latin typeface="Calibri"/>
                  </a:rPr>
                  <a:t xml:space="default">apricot</a:t>
                </a:r>
                <a:endParaRPr sz="1000" lang="en-US"/>
              </a:p>
            </a:txBody>
            <a:tcPr marT="45720" marB="45720" marL="45720" marR="45720" anchor="ctr">
              <a:tcBdr>
                <a:lnT w="4762" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="4762" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnB>
                <a:lnL w="4762" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnL>
                <a:lnR w="4762" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnR>
              </a:tcBdr>
              <a:noFill/>
            </a:tcPr>
          </a:tc>
          <a:tc>
            <a:txBody>
              <a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" rtlCol="0" anchor="ctr"/>
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
                  <a:rPr lang="en-US" sz="1000">
                    <a:latin typeface="Calibri"/>
                  </a:rPr>
                  <a:t xml:space="default">one</a:t>
                </a:r>
                <a:endParaRPr sz="1000" lang="en-US"/>
              </a:p>
            </a:txBody>
            <a:tcPr marT="45720" marB="45720" marL="45720" marR="45720" anchor="ctr">
              <a:tcBdr>
                <a:lnT w="4762" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="4762" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnB>
                <a:lnL w="4762" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnL>
                <a:lnR w="4762" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnR>
              </a:tcBdr>
              <a:noFill/>
            </a:tcPr>
          </a:tc>
          <a:tc>
            <a:txBody>
              <a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" rtlCol="0" anchor="ctr"/>
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
                  <a:rPr lang="en-US" sz="1000">
                    <a:latin typeface="Calibri"/>
                  </a:rPr>
                  <a:t xml:space="default">2015-01-15</a:t>
                </a:r>
                <a:endParaRPr sz="1000" lang="en-US"/>
              </a:p>
            </a:txBody>
            <a:tcPr marT="45720" marB="45720" marL="45720" marR="45720" anchor="ctr">
              <a:tcBdr>
                <a:lnT w="4762" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="4762" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnB>
                <a:lnL w="4762" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnL>
                <a:lnR w="4762" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnR>
              </a:tcBdr>
              <a:noFill/>
            </a:tcPr>
          </a:tc>
          <a:tc>
            <a:txBody>
              <a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" rtlCol="0" anchor="ctr"/>
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
                  <a:rPr lang="en-US" sz="1000">
                    <a:latin typeface="Calibri"/>
                  </a:rPr>
                  <a:t xml:space="default">13:35</a:t>
                </a:r>
                <a:endParaRPr sz="1000" lang="en-US"/>
              </a:p>
            </a:txBody>
            <a:tcPr marT="45720" marB="45720" marL="45720" marR="45720" anchor="ctr">
              <a:tcBdr>
                <a:lnT w="4762" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="4762" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnB>
                <a:lnL w="4762" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnL>
                <a:lnR w="4762" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnR>
              </a:tcBdr>
              <a:noFill/>
            </a:tcPr>
          </a:tc>
          <a:tc>
            <a:txBody>
              <a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" rtlCol="0" anchor="ctr"/>
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
                  <a:rPr lang="en-US" sz="1000">
                    <a:latin typeface="Calibri"/>
                  </a:rPr>
                  <a:t xml:space="default">2018-01-01 02:22</a:t>
                </a:r>
                <a:endParaRPr sz="1000" lang="en-US"/>
              </a:p>
            </a:txBody>
            <a:tcPr marT="45720" marB="45720" marL="45720" marR="45720" anchor="ctr">
              <a:tcBdr>
                <a:lnT w="4762" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="4762" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnB>
                <a:lnL w="4762" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnL>
                <a:lnR w="4762" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnR>
              </a:tcBdr>
              <a:noFill/>
            </a:tcPr>
          </a:tc>
          <a:tc>
            <a:txBody>
              <a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" rtlCol="0" anchor="ctr"/>
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
                  <a:rPr lang="en-US" sz="1000">
                    <a:latin typeface="Calibri"/>
                  </a:rPr>
                  <a:t xml:space="default">49.95</a:t>
                </a:r>
                <a:endParaRPr sz="1000" lang="en-US"/>
              </a:p>
            </a:txBody>
            <a:tcPr marT="45720" marB="45720" marL="45720" marR="45720" anchor="ctr">
              <a:tcBdr>
                <a:lnT w="4762" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="4762" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnB>
                <a:lnL w="4762" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnL>
                <a:lnR w="4762" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnR>
              </a:tcBdr>
              <a:noFill/>
            </a:tcPr>
          </a:tc>
          <a:tc>
            <a:txBody>
              <a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" rtlCol="0" anchor="ctr"/>
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
                  <a:rPr lang="en-US" sz="1000">
                    <a:latin typeface="Calibri"/>
                  </a:rPr>
                  <a:t xml:space="default">row_1</a:t>
                </a:r>
                <a:endParaRPr sz="1000" lang="en-US"/>
              </a:p>
            </a:txBody>
            <a:tcPr marT="45720" marB="45720" marL="45720" marR="45720" anchor="ctr">
              <a:tcBdr>
                <a:lnT w="4762" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="4762" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnB>
                <a:lnL w="4762" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnL>
                <a:lnR w="4762" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnR>
              </a:tcBdr>
              <a:noFill/>
            </a:tcPr>
          </a:tc>
          <a:tc>
            <a:txBody>
              <a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" rtlCol="0" anchor="ctr"/>
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
                  <a:rPr lang="en-US" sz="1000">
                    <a:latin typeface="Calibri"/>
                  </a:rPr>
                  <a:t xml:space="default">grp_a</a:t>
                </a:r>
                <a:endParaRPr sz="1000" lang="en-US"/>
              </a:p>
            </a:txBody>
            <a:tcPr marT="45720" marB="45720" marL="45720" marR="45720" anchor="ctr">
              <a:tcBdr>
                <a:lnT w="4762" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnT>
                <a:lnB w="4762" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnB>
                <a:lnL w="4762" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnL>
                <a:lnR w="4762" cap="flat" cmpd="single" algn="ctr">
                  <a:solidFill>
                    <a:srgbClr val="D3D3D3"/>
                  </a:solidFill>
                </a:lnR>
              </a:tcBdr>
              <a:noFill/>
            </a:tcPr>
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
          <a:rPr lang="en-US" sz="1200">
            <a:latin typeface="Calibri"/>
            <a:solidFill>
              <a:srgbClr val="333333"/>
            </a:solidFill>
          </a:rPr>
          <a:t xml:space="default">TABLE TITLE</a:t>
        </a:r>
        <a:endParaRPr sz="1000" lang="en-US"/>
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
          <a:rPr lang="en-US" sz="800">
            <a:latin typeface="Calibri"/>
            <a:solidFill>
              <a:srgbClr val="333333"/>
            </a:solidFill>
          </a:rPr>
          <a:t xml:space="default">table subtitle</a:t>
        </a:r>
        <a:endParaRPr sz="1000" lang="en-US"/>
      </a:p>

