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

