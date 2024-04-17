from _typeshed import Incomplete, SupportsWrite
from collections.abc import Callable, Mapping, MutableMapping
from typing import Any
from typing_extensions import TypeAlias
from xml.sax import handler
from xml.sax.xmlreader import AttributesImpl, XMLReader

from netaddr.core import Publisher, Subscriber
from netaddr.ip import IPAddress, IPNetwork, IPRange

_IanaInfoKey: TypeAlias = IPAddress | IPNetwork | IPRange

IANA_INFO: dict[str, dict[_IanaInfoKey, dict[str, str]]]

class SaxRecordParser(handler.ContentHandler):
    def __init__(self, callback: Callable[[Mapping[str, object] | None], object] | None = None) -> None: ...
    def startElement(self, name: str, attrs: AttributesImpl) -> None: ...
    def endElement(self, name: str) -> None: ...
    def characters(self, content: str) -> None: ...

class XMLRecordParser(Publisher):
    xmlparser: XMLReader
    fh: Incomplete
    def __init__(self, fh: Incomplete, **kwargs: object) -> None: ...
    def process_record(self, rec: Mapping[str, object]) -> dict[str, str] | None: ...
    def consume_record(self, rec: object) -> None: ...
    def parse(self) -> None: ...
    # Arbitrary attributes are set in __init__ with `self.__dict__.update(kwargs)`
    def __getattr__(self, name: str, /) -> Any: ...

class IPv4Parser(XMLRecordParser):
    def process_record(self, rec: Mapping[str, object]) -> dict[str, str]: ...

class IPv6Parser(XMLRecordParser):
    def process_record(self, rec: Mapping[str, object]) -> dict[str, str]: ...

class IPv6UnicastParser(XMLRecordParser):
    def process_record(self, rec: Mapping[str, object]) -> dict[str, str]: ...

class MulticastParser(XMLRecordParser):
    def normalise_addr(self, addr: str) -> str: ...

class DictUpdater(Subscriber):
    dct: MutableMapping[_IanaInfoKey, Incomplete]
    topic: str
    unique_key: str
    def __init__(self, dct: MutableMapping[_IanaInfoKey, Incomplete], topic: str, unique_key: str) -> None: ...
    def update(self, data: Incomplete) -> None: ...

def load_info() -> None: ...
def pprint_info(fh: SupportsWrite[str] | None = None) -> None: ...
def query(ip_addr: IPAddress) -> dict[str, list[dict[str, str]]]: ...
