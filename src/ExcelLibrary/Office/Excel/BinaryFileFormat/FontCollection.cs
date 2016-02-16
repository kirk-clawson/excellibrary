using System.Collections.Generic;
using System.Linq;
using ExcelLibrary.SpreadSheet;

namespace ExcelLibrary.BinaryFileFormat
{
    public class FontCollection
    {
        private readonly IList<FONT> _fonts = new List<FONT>();
        private readonly IDictionary<string, ushort> _indexLookup = new Dictionary<string, ushort>();

        public FONT GetFont(ushort index)
        {
            if (index <= 3) return _fonts[index];
            if (index > 4) return _fonts[index - 1];
            throw new KeyNotFoundException("Index value of 4 is invalid for FONT collection");
        }

        public ushort GetIndex(CellFont font)
        {
            string fontKey = font.ToString();
            if (_indexLookup.ContainsKey(fontKey))
            {
                return _indexLookup[fontKey];
            }

            var f = new FONT
            {
                Height = font.Height,
                ColorIndex = font.Color.Index,
                Escapement = (ushort) font.Escapement,
                UnderlineStyle = (byte) font.UnderlineStyle,
                Family = (byte) font.Family,
                CharacterSet = 1,
                Name = font.Name,
                Bold = font.Bold,
                Italic = font.Italic,
                Strikethrough = font.Strikethrough,
                Outlined = font.Outlined,
                Shadowed = font.Shadowed,
                Condensed = font.Condensed,
                Extended = font.Extended
            };

            _fonts.Add(f);
            _indexLookup.Add(fontKey, (ushort) (_fonts.Count - 1));
            return (ushort) (_fonts.Count - 1);
        }

        public void Add(FONT font)
        {
            _fonts.Add(font);
        }

        public FONT[] ToArray()
        {
            return _fonts.ToArray();
        }
    }
}