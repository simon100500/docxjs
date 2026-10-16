describe('Agen fidelity regressions', function () {
  const xml = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:rPr><w:rFonts w:ascii="Times New Roman"/><w:sz w:val="18"/></w:rPr><w:tabs><w:tab w:val="right" w:pos="4800"/></w:tabs></w:pPr><w:bookmarkStart w:id="1" w:name="field"/><w:r><w:rPr><w:sz w:val="18"/><w:u w:val="single"/></w:rPr><w:t>Text</w:t><w:tab/></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>';
  async function documentBytes() {
    const zip = new JSZip();
    zip.file('[Content_Types].xml','<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
    zip.file('_rels/.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
    zip.file('word/document.xml',xml);
    return zip.generateAsync({type:'uint8array'});
  }
  for(const zoom of [1,.5,1.25]) it(`lays out trailing underlined tabs before resolving at zoom ${zoom}`,async function(){
    const body=document.createElement('div'),styles=document.createElement('div');
    body.style.zoom=String(zoom);document.body.append(styles,body);
    try {
      await docx.renderAsync(await documentBytes(),body,styles,{experimental:true});
      const p=body.querySelector('p'),tab=body.querySelector('.docx-tab-stop');
      expect(p.style.fontSize).toBe('9pt');
      expect(getComputedStyle(body.querySelector('#field')).fontSize).toBe('0px');
      expect(tab.style.display).toBe('inline-block');
      const endpoint=(tab.getBoundingClientRect().right-p.getBoundingClientRect().left)/zoom;
      expect(Math.abs(endpoint-320)).toBeLessThan(2);
      expect(parseFloat(tab.style.width)).toBeGreaterThan(0);
    } finally {body.remove();styles.remove();}
  });
  it('preserves authoritative pages and returns the original PDF bytes', async function(){
    const body=document.createElement('div'),styles=document.createElement('div');
    const bytes=new Uint8Array([37,80,68,70,45]);
    const imageUrl='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aX1cAAAAASUVORK5CYII=';
    const result=await docx.renderAsync(new Uint8Array(),body,styles,{renderMode:'pdf',pdfAdapter:{convert:async()=>bytes,renderPages:async(pdf)=>{pdf[0]=0;return [{width:794,height:1123,imageUrl},{width:794,height:1123,imageUrl}];}}});
    expect(body.dataset.renderMode).toBe('pdf');expect(body.querySelectorAll('section').length).toBe(2);
    expect(result.pageCount).toBe(2);expect(result.pdfBytes).toEqual(bytes);expect(result.pdfBytes[0]).toBe(37);
  });
  it('does not silently approximate an exact preview when the adapter fails',async function(){
    const body=document.createElement('div');
    await expectAsync(docx.renderAsync(new Uint8Array(),body,null,{renderMode:'pdf'})).toBeRejectedWithError(/authoritative PDF adapter/);
    expect(body.childNodes.length).toBe(0);
  });
});
