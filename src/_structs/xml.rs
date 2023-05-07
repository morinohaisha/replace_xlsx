use quick_xml::{events::Event, reader::Reader};

pub struct XmlReader<'a> {
    pub reader: Reader<&'a [u8]>,
}

impl XmlReader<'_> {
    pub fn new(xml: &str) -> XmlReader {
        let reader = Reader::from_str(xml);
        XmlReader { reader }
    }
    pub fn trim_text(&mut self, value: bool) {
        self.reader.trim_text(value);
    }

    pub fn read_event_into<'b>(&mut self, buf: &'b mut Vec<u8>) -> anyhow::Result<Event<'b>> {
        Ok(self.reader.read_event_into(buf)?)
    }

    pub fn buffer_position(&self) -> usize {
        self.reader.buffer_position()
    }
}
