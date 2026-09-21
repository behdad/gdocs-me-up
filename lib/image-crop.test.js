const { imageCrop, cropImageStyle } = require('./image-crop');

test('ignores absent crops and floating-point residue', () => {
  for (const crop of [undefined, null, {}, { offsetLeft: 1e-16, offsetRight: -1e-16 }]) {
    expect(imageCrop(crop)).toBeNull();
  }
});

test('uses all four edges, including crops extending outside the image', () => {
  const crop = imageCrop({ offsetLeft: .1, offsetRight: .3, offsetTop: .2, offsetBottom: .4 });
  expect(crop.width).toBeCloseTo(.6);
  expect(crop.height).toBeCloseTo(.4);
  expect(imageCrop({ offsetLeft: -.5, offsetRight: -.5 }).width).toBe(2);
});

test('rejects empty or invalid crop rectangles', () => {
  expect(() => imageCrop({ offsetLeft: .5, offsetRight: .5 })).toThrow('Empty');
  expect(() => imageCrop({ angle: NaN })).toThrow('Invalid');
});

test('maps an asymmetric crop to the frame', () => {
  const css = cropImageStyle(imageCrop({ offsetLeft: .125, offsetRight: .375, offsetBottom: .5 }), {});
  expect(css).toContain('width:200%;height:200%;left:-25%;top:0%;');
});

test('undoes clockwise rotation around the crop centre with unequal scaling', () => {
  const css = cropImageStyle(imageCrop({ angle: Math.PI / 2 }), {
    displayWidth: 200, displayHeight: 100, sourceWidth: 100, sourceHeight: 100
  });
  expect(css).toContain('transform-origin:50% 50%;transform:matrix(0,-0.5,2,0,0,0);');
});
