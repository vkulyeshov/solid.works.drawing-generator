"""Rebuild an existing VBA compound-file container with reviewed module source.

Preserves references and directory metadata. Removes compiled performance caches;
SolidWorks must compile the supplied VBA source on first load.
No third-party Python packages required.
"""
from pathlib import Path
import struct, math, sys
sys.dont_write_bytecode = True

FREE, END, FATSECT = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD
u16 = lambda b, p: struct.unpack_from('<H', b, p)[0]
u32 = lambda b, p: struct.unpack_from('<I', b, p)[0]

class Compound:
    def __init__(self, path):
        self.raw = Path(path).read_bytes()
        assert self.raw[:8] == bytes.fromhex('d0cf11e0a1b11ae1')
        self.size = 1 << u16(self.raw, 30)
        ids = list(struct.unpack_from('<109I', self.raw, 76))
        difat = u32(self.raw, 68)
        while difat not in (END, FREE):
            values = struct.unpack('<%dI' % (self.size//4), self.sector(difat))
            ids += list(values[:-1]); difat = values[-1]
        self.fat = []
        for n in ids[:u32(self.raw, 44)]:
            self.fat += list(struct.unpack('<%dI' % (self.size//4), self.sector(n)))
        directory = self.chain(u32(self.raw, 48))
        self.entries = [bytearray(directory[i:i+128]) for i in range(0,len(directory),128)]
        root = self.entries[0]
        self.mini = self.chain(u32(root,116))[:struct.unpack_from('<Q',root,120)[0]]
        mini_fat_data = self.chain(u32(self.raw,60))
        self.minifat = list(struct.unpack('<%dI' % (len(mini_fat_data)//4),mini_fat_data))
        self.streams={}
        for i,e in enumerate(self.entries):
            if e[66] != 2: continue
            length=struct.unpack_from('<Q',e,120)[0]
            first=u32(e,116)
            self.streams[i] = (self.chain(first,True) if length < 4096 else self.chain(first))[:length]
    def sector(self,i): return self.raw[(i+1)*self.size:(i+2)*self.size]
    def chain(self,first,mini=False):
        values=[]; seen=set(); table=self.minifat if mini else self.fat
        while first not in (END,FREE):
            assert first not in seen; seen.add(first)
            values.append(self.mini[first*64:(first+1)*64] if mini else self.sector(first))
            first=table[first]
        return b''.join(values)
    def name(self,i):
        e=self.entries[i]; return bytes(e[:u16(e,64)-2]).decode('utf-16le') if u16(e,64)>0 else ''
    def write(self,path):
        entries=[bytearray(e) for e in self.entries]
        sectors=[]; fat=[]; mini=[]; minifat=[]
        def alloc(data):
            if not data: return END
            start=len(sectors); count=math.ceil(len(data)/512)
            for n in range(count):
                sectors.append(data[n*512:(n+1)*512].ljust(512,b'\0'))
                fat.append(start+n+1 if n<count-1 else END)
            return start
        for i,data in self.streams.items():
            if len(data)<4096:
                start=len(mini) if data else END
                count=math.ceil(len(data)/64)
                for n in range(count):
                    mini.append(data[n*64:(n+1)*64].ljust(64,b'\0'))
                    minifat.append(start+n+1 if n<count-1 else END)
            else: start=alloc(data)
            struct.pack_into('<IQ',entries[i],116,start,len(data))
        rootmini=b''.join(mini)
        struct.pack_into('<IQ',entries[0],116,alloc(rootmini),len(rootmini))
        minifatdata=struct.pack('<%dI'%len(minifat),*minifat)
        minifatfirst=alloc(minifatdata)
        directory=b''.join(entries)
        dirfirst=alloc(directory)
        fatcount=1
        while fatcount*128 < len(sectors)+fatcount: fatcount+=1
        assert fatcount<=109
        fatids=list(range(len(sectors),len(sectors)+fatcount))
        fat += [FATSECT]*fatcount
        fat += [FREE]*(fatcount*128-len(fat))
        fatbytes=struct.pack('<%dI'%len(fat),*fat)
        sectors += [fatbytes[i*512:(i+1)*512] for i in range(fatcount)]
        h=bytearray(self.raw[:512])
        struct.pack_into('<HH',h,26,3,0xFFFE)
        struct.pack_into('<HH',h,30,9,6)
        struct.pack_into('<IIIIIIIII',h,40,0,fatcount,dirfirst,0,4096,minifatfirst,math.ceil(len(minifatdata)/512),END,0)
        struct.pack_into('<109I',h,76,*(fatids+[FREE]*(109-fatcount)))
        Path(path).write_bytes(h+b''.join(sectors))

def decompress(data):
    assert data[0]==1
    result=bytearray(); pos=1
    while pos+2<=len(data):
        header=u16(data,pos); size=(header&4095)+3
        chunk=data[pos+2:pos+size]; pos+=size
        if not header&0x8000: result+=chunk; continue
        out=bytearray(); p=0
        while p<len(chunk):
            flags=chunk[p];p+=1
            for bit in range(8):
                if p>=len(chunk): break
                if flags&(1<<bit):
                    token=u16(chunk,p);p+=2
                    bits=max(4,(len(out)-1).bit_length())
                    length=(token&((1<<(16-bits))-1))+3
                    offset=(token>>(16-bits))+1
                    for _ in range(length): out.append(out[-offset])
                else: out.append(chunk[p]);p+=1
        result+=out
    return bytes(result)

def compress(data):
    # Valid MS-OVBA compressed chunks containing literal tokens only.
    result=bytearray(b'\x01')
    for start in range(0,len(data),3600):
        plain=data[start:start+3600]
        encoded=b''.join(b'\0'+plain[i:i+8] for i in range(0,len(plain),8))
        result+=struct.pack('<H',0xB000|(len(encoded)-1))+encoded
    assert decompress(result)==data
    return bytes(result)

if __name__=='__main__':
    source=Path(sys.argv[1]); c=Compound(source)
    if len(sys.argv)==2:
        for i,data in c.streams.items(): print(i,repr(c.name(i)),len(data))
    else:
        vbadir=next(i for i in c.streams if c.name(i)=='dir')
        d=bytearray(decompress(c.streams[vbadir]))
        pattern=bytes.fromhex('310004000000')
        positions=[];p=0
        while True:
            p=d.find(pattern,p)
            if p<0:break
            positions.append(p);p+=len(pattern)
        assert len(positions)==2,positions
        p=positions[-1]+6
        old_offset=u32(d,p)
        project=c.streams[next(i for i in c.streams if c.name(i)=='PROJECT')].decode('cp1252')
        import re
        module_name=next(x[7:] for x in project.splitlines() if x.startswith("Module="))
        module_id=next(i for i in c.streams if c.name(i)==module_name)
        oldsource=decompress(c.streams[module_id][old_offset:])
        print('Container source:',oldsource[:80])
        # Keep the module name stored in the container's dir and PROJECT streams.
        newsource=Path(sys.argv[2]).read_text('utf-8-sig').replace('Attribute VB_Name = "generator"','Attribute VB_Name = "'+module_name+'"').replace('\r\n','\n').replace('\n','\r\n').encode('cp1251')
        struct.pack_into('<I',d,p,0)
        libid=next(i for i in c.streams if c.name(i)=='ThisLibrary')
        liboffset=u32(d,positions[0]+6)
        c.streams[libid]=compress(decompress(c.streams[libid][liboffset:]))
        struct.pack_into('<I',d,positions[0]+6,0)
        refstart=d.find(bytes.fromhex('160007000000')+b'SwConst')
        refend=d.find(bytes.fromhex('0f00020000000200'),refstart)
        if refstart>0 and refend>refstart: d=d[:refstart]+d[refend:]
        cp=d.find(bytes.fromhex('030002000000'))+6
        struct.pack_into('<H',d,cp,1251)
        c.streams[vbadir]=compress(d)
        c.streams[module_id]=compress(newsource)
        for i in c.streams:
            if c.name(i)=='_VBA_PROJECT':
                # Version-independent header and no performance cache (MS-OVBA 2.3.4.1).
                c.streams[i]=bytes.fromhex('cc61ffff000000')
            elif c.name(i).startswith('__SRP_'): c.streams[i]=b''
        from project_paths import external_path
        c.write(external_path(sys.argv[3]))
        check=Compound(sys.argv[3])
        assert decompress(check.streams[module_id])==newsource
        print('Wrote',sys.argv[3])




