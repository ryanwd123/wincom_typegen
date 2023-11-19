# %%
from pathlib import Path
import win32com.client as win
from win32com.client.selecttlb import EnumTlbs

# import win32typing as wint
# import win32helper.win32typing as wint
import  wincom_typegen.typegen_classes as wint
# import _win32typing
# from types
from pythoncom import LoadTypeLib, CreateBindCtx, GetRunningObjectTable
from dataclasses import dataclass
import pprint
from enum import Enum
import ctypes.wintypes
import comtypes
import ctypes


def get_type_info(obj) -> wint.PyITypeInfo:
    return getattr(obj, "_oleobj_", obj).GetTypeInfo()

def get_type_name(ojb) -> str:
    return get_type_info(ojb).GetDocumentation(-1)[0]

def get_type_lib(obj) -> wint.PyITypeLib:
    return get_type_info(obj).GetContainingTypeLib()[0]


target_objects = [
    ("Word", "Application"),
    ("Word", "Document"),
    ("Excel", "Application"),
    ("Excel", "Chart"),
    ("PowerPoint", "Application"),
    ("Access", "Application"),
    ("Publisher", "Document"),
    ("Publisher", "Application"),
    ("Outlook", "Application"),
    ("Scripting", "Dictionary"),
    ("Scripting", "FileSystemObject"),
    ("Scripting", "Encoder"),
    ("ADODB", "Connection"),
    ("ADODB", "Record"),
    ("ADODB", "Stream"),
    ("ADODB", "Command"),
    ("ADODB", "Recordset"),
    ("ADODB", "Parameter"),
    ("ADOX", "Table"),
    ("ADOX", "Column"),
    ("ADOX", "Index"),
    ("ADOX", "Key"),
    ("ADOX", "Group"),
    ("ADOX", "User"),
    ("ADOX", "Catalog"),
    ("ADOR", "Recordset"),
    ("ADOMD", "Catalog"),
    ("ADOMD", "Cellset"),
    ("ADOMD", "Cellset"),
]


@dataclass
class ImplType:
    id: int
    flag: int
    ref: str
    top_id: int
    name: str


@dataclass
class TypeAttrs:
    cFuncs: int
    cImplTypes: int
    cVars: int
    cbAlignment: int
    cbSizeInstance: int
    cbSizeVft: int
    idldescType: tuple
    iid: wint.PyIID
    lcid: int
    memidConstructor: int
    memidDestructor: int
    tdescAlias: str
    typekind: int
    wMajorVerNum: int
    wMinorVerNum: int
    wTypeFlags: int


class RetType(Enum):
    LONG = 3
    SINGLE = 4
    DOUBLE = 5
    DATE = 7
    STRING = 8
    OBJECT = 9
    BOOL = 11
    VARIANT = 12
    LONG2 = 22
    NOTHING = 24
    CLASS = 26
    VARIANT_ARRAY = 27
    ENUM = 29


ret_type_map = {
    2: "int",
    3: "int",
    4: "int",
    5: "float",
    6: "float",
    7: "datetime.date",
    8: "str",
    9: "VbObject",
    12: "VbVariant",
    25: "VbVariant",
    13: "VbUnknown",
    11: "bool",
    22: "int",
    21: "int",
    17: "bytes",
    24: "x",
}


constants_tdesc_alias_map = {
    3: ctypes.c_int,
    4: ctypes.c_float,
    6: ctypes.c_longlong,
    19: ctypes.c_ulong,
    8: comtypes.BSTR,
    11: ctypes.wintypes.VARIANT_BOOL,
    20: ctypes.c_long,
    21: ctypes.c_ulong,
    22: ctypes.c_int,
}


@dataclass
class ComArg:
    id: int
    name: str
    type: "ComRefType"
    optional: bool
    code: tuple
    parent_lib: str

    def pyi_str(self, optional=False):
        if not self.parent_lib:
            print(self.name)
        try:
            qq = self.type.pyi_str(self.parent_lib)
            typ = f":'{qq}'" if self.type else ""
        except Exception as e:
            # print(self.name, self.parent_lib, self.type, self.code, self.type.name, self.type.lib_name)
            # raise e
            typ = ""

        if optional:
            opt = True
        else:
            opt = self.optional

        return f'{self.name}{typ}{" = None" if opt else ""}'


@dataclass
class ComFunc:
    id: int
    desc: wint.FUNCDESC
    names: tuple[str]
    doc: tuple
    parent_name: str
    ret_type: "ComRefType"
    args: list[ComArg]
    lib_name: str

    # def __repr__(self) -> str:
    # return f'id:{self.id: <3} kind:{self.desc.funckind}  inv:{self.desc.invkind} flags:{self.desc.wFuncFlags} args_count:{len(self.args)}  name:{self.parent_name}.{self.names[0]}   ret:{self.ret_type}'


@dataclass
class ComVar:
    id: int
    name: str
    value: int
    desc: wint.VARDESC
    doc: tuple
    parent_name: str

    def pyi_str(self, py=False):
        if py:
            if isinstance(self.value, int):
                return f"    {self.name} = {self.value}"
            else:
                return f'    {self.name} = "{self.value}"'
        else:
            if not self.value:
                return f"    {self.name}:VbVariant"
            return f"    {self.name}:{type(self.value).__name__}"

    def __repr__(self) -> str:
        return f"{self.name} = {self.value}"


@dataclass
class ComRefType:
    name: str
    lib_name: str
    type: wint.PyITypeInfo

    def pyi_str(self, other):
        if self.lib_name and self.lib_name != other:
            return f"l_{self.lib_name.lower()}.{self.name}"
        else:
            return self.name


@dataclass
class ComFunction:
    funcs: list[ComFunc]
    name: str
    ret_type: ComRefType
    args: list[ComArg]
    prop: bool
    read_only: bool
    parent: str
    lib_name: str
    get_flag: bool
    set_flag: bool

    @property
    def pyi_str(self):
        txt = ""

        # if self.prop and len(self.args) == 0:
        if self.set_flag:
            txt += "    @property\n"

        args = ""
        if len(self.args) > 0:
            if (
                len(self.args) > 10
                and self.args[0].name.startswith("Arg")
                and self.args[-1].name.startswith("Arg")
            ):
                args = "self, *args"
            else:
                opt = False
                arg_list = ["self"]
                for arg in self.args:
                    if arg.optional:
                        opt = True
                    if arg.name == "RHS":
                        continue
                    arg_list.append(arg.pyi_str(opt))
                args = ", ".join(arg_list)

        ret = ""
        if self.ret_type:
            ret = f" -> '{self.ret_type.pyi_str(self.lib_name)}'"

        if self.prop and len(self.args) == 0:
            if self.ret_type:
                ret = f":'{self.ret_type.pyi_str(self.lib_name)}'"
            else:
                ret = ":VbVariant"
            return f"    {self.name}{ret}"

        txt += f"    def {self.name}({args}){ret}: ..."
        return txt


@dataclass
class ComObject:
    id: int
    name: str
    info: wint.PyITypeInfo
    impl_types: list["ComObject"]
    doc: tuple
    attrs: TypeAttrs
    funcs: list[ComFunc]
    funcs_map: dict[str, list[ComFunc]]
    vars: list[ComVar]
    lib_name: str

    def __post_init__(self):
        all_funcs: dict[str, list[ComFunc]] = {}
        all_funcs.update(self.funcs_map)
        for obj in self.impl_types:
            if not obj:
                print(f"{self.lib_name}.{self.name} has None type")
                continue
            # all_funcs.update(obj.funcs_map)

        self.funcs2: dict[str, ComFunction] = {}

        for k, v in all_funcs.items():
            _get = False
            _set = False

            read_only = True
            property = True
            ret = None
            args: list[ComArg] = []

            for f in v:
                if f.desc.invkind == 2:
                    _get = True
                if f.desc.invkind == 1:
                    property = False
                if f.ret_type:
                    ret = f.ret_type
                if f.desc.invkind == 4:
                    _set = True
                    read_only = False
                    continue
                args = f.args

            self.funcs2[k] = ComFunction(
                v,
                k,
                ret,
                args,
                property,
                read_only,
                f.parent_name,
                f.lib_name,
                _get,
                _set,
            )

    def pyi_str(self, ref_types, object_list: list[str], py=False):
        txt = ""
        kind = self.attrs.typekind
        # if self.name[0] == '_':
        # return ''
        if kind == 6:
            if py:
                return ""
            t_alias = self.attrs.tdescAlias
            if isinstance(t_alias, tuple):
                ref = get_ref_type_name(
                    (t_alias,), self.info, self.name, self.lib_name, "", ref_types
                )
                if not ref:
                    unkown_types.append([t_alias, self.lib_name, self.name])
                    # print(f'{self.lib_name}.{self.name} has unknown type {t_alias}')
                    return ""

                txt = f"{self.name}:'{ref.name}' = None\n"
                return txt

            if t_alias not in constants_tdesc_alias_map:
                print(f"{self.lib_name}.{self.name} has unknown type {t_alias}")
                pprint.pprint(self)
                return ""
            t = constants_tdesc_alias_map[t_alias]
            txt = f"{self.name} = {t.__module__}.{t.__name__}\n"

            return txt
        if len(self.vars) > 0:
            txt = f"class {self.name}:\n"
            txt += "\n".join([var.pyi_str(py) for var in self.vars])
            txt += "\n\n"
        elif len(self.funcs2) > 0 or self.impl_types:
            if py:
                return f"class {self.name}(ComDummy):\n    pass\n\n"
            iter_txt = ""
            if "Item" in self.funcs_map:
                ret = self.funcs_map["Item"][0].ret_type
                if ret:
                    iter_txt = f"    def __iter__(self) -> '{self.name}': ...\n"
                    iter_txt += f"    def __next__(self) -> '{ret.pyi_str(ret.lib_name)}': ...\n"
                    iter_txt += f"    def __call__(self, index:int) -> '{ret.pyi_str(ret.lib_name)}': ...\n"
                    iter_txt += f"    def __getitem__(self, items) -> '{ret.pyi_str(ret.lib_name)}': ...\n"

            imps = ""
            if self.impl_types:
                imps = ", ".join(
                    [
                        f"{i.name}"
                        for i in self.impl_types
                        if i.name in object_list + ["ComDummy"]
                    ]
                )
                imps = f"({imps})"

            txt = f"class {self.name}{imps}:\n"
            txt += iter_txt
            funcs = "\n".join([f.pyi_str for f in self.funcs2.values()])
            txt += funcs
            if len(funcs) == 0:
                txt += "    pass\n"
            txt += "\n\n"
        else:
            txt = f"class {self.name}:\n    pass\n"
        return "\n" + txt


# '\n'.join([])
# func_ret_types = set()
# sample_funcs:list[ComFunc] = []
# sample_args = []
# sample_args_t = set()


def get_first_int(t: tuple):
    if type(t[0]) == int:
        return t[0]
    else:
        get_first_int(t[0])


unkown_types = []


def get_ref_type_name(
    refttype: tuple, info: wint.PyITypeInfo, name, fn_name, lib_name, ref_types
):
    if isinstance(refttype[0], int):
        ret = ret_type_map.get(refttype[0], None)
        if ret == "x":
            return
        if not ret:
            unkown_types.append([lib_name, fn_name, name, refttype])
            # print(f'{lib_name}.{fn_name}.{name} unknown type {refttype}')
            return None
        return ComRefType(ret, None, None)
    else:
        ret: str = None
        if refttype[0][0] == 26:
            typ = info.GetRefTypeInfo(refttype[0][1][1])
            ret = typ.GetDocumentation(-1)[0]
        if refttype[0][0] == 29:
            typ = info.GetRefTypeInfo(refttype[0][1])
            ret = typ.GetDocumentation(-1)[0]
        if not ret:
            return None
        lib = typ.GetContainingTypeLib()[0]
        lib_name = lib.GetDocumentation(-1)[0]

        # ret = ret.lstrip('_')
        if not ret:
            return None

        ref_types[ret] = typ
        return ComRefType(ret, lib_name, typ)


def get_vars(info: wint.PyITypeInfo, attrs: TypeAttrs, parent_name: str):
    vars = []
    for i in range(attrs.cVars):
        desc: wint.VARDESC = info.GetVarDesc(i)
        doc = info.GetDocumentation(desc.memid)
        vars.append(ComVar(i, doc[0], desc.value, desc, doc, parent_name))
    return vars


def get_funcs(
    info: wint.PyITypeInfo, attrs: TypeAttrs, parent_name: str, lib_name: str, ref_types
):
    funcs = []
    funcs_map: dict[str, list[ComFunc]] = {}
    for i in range(attrs.cFuncs):
        desc: wint.FUNCDESC = info.GetFuncDesc(i)
        if desc.wFuncFlags == 1:
            continue
        names = info.GetNames(desc.memid)
        name = names[0]
        if name[0] == "_":
            continue
        doc = info.GetDocumentation(desc.memid)
        ret = get_ref_type_name(
            desc.rettype, info, name, parent_name, lib_name, ref_types
        )
        args = []
        for i, arg in enumerate(desc.args):
            optional = False
            arg_t = get_first_int(arg)
            try:
                arg_name = names[i + 1]
            except:
                arg_name = "unknown"
            try:
                arg_type = get_ref_type_name(
                    arg, info, name, parent_name, lib_name, ref_types
                )
            except:
                # print(f'{parent_name}.{arg_name}', names)
                arg_type = None
            if arg[1] in (17, 49):
                optional = True

            # print(f'{parent_name}.{arg_name}')
            args.append(ComArg(i, arg_name, arg_type, optional, arg, lib_name))

        fnc = ComFunc(i, desc, names, doc, parent_name, ret, args, lib_name)
        funcs.append(fnc)
        if name in funcs_map:
            funcs_map[name].append(fnc)
        else:
            funcs_map[name] = [fnc]

        # ret_type_id = get_first_int(desc.rettype)
        # if ret_type_id not in func_ret_types:
        #     func_ret_types.add(ret_type_id)
        #     sample_funcs.append(fnc)

    return funcs, funcs_map


def get_impl_types(
    tl: wint.PyITypeLib,
    info: wint.PyITypeInfo,
    impl_types_cnt,
    lib_name: str,
    ref_types,
):
    impl_types: list[ComObject] = []
    for i in range(impl_types_cnt):
        flag = info.GetImplTypeFlags(i)
        if flag == 0:
            continue
        ref = info.GetRefTypeOfImplType(i)
        top_id = info.GetRefTypeInfo(ref).GetContainingTypeLib()[1]
        name = tl.GetDocumentation(top_id)[0]
        ojb = get_obj(top_id, tl, lib_name, ref_types)
        impl_types.append(ojb)
        # impl_types.append(ImplType(
        #     id=i,
        #     flag=flag,
        #     ref=ref,
        #     top_id=top_id,
        #     name=name

        #     ))
    return impl_types


def get_attrs(at: wint.TYPEATTR):
    attr_dict = {}
    for attr in dir(at):
        if attr[0] == "_":
            continue
        attr_dict[attr] = getattr(at, attr)
        # print(f'{attr}: {type(getattr(at, attr)).__name__}')
    return TypeAttrs(**attr_dict)


def get_obj(i, tl: wint.PyITypeLib, lib_name: str, ref_types):
    info = tl.GetTypeInfo(i)
    doc = tl.GetDocumentation(i)
    attrs = info.GetTypeAttr()
    attrs.cImplTypes
    # if attrs.typekind in [3]:

    # return
    impl_types = get_impl_types(tl, info, attrs.cImplTypes, lib_name, ref_types)
    attrs_obj = get_attrs(attrs)
    funcs, funcs_map = get_funcs(info, attrs_obj, doc[0], lib_name, ref_types)

    obj = ComObject(
        id=i,
        name=doc[0],
        info=info,
        impl_types=impl_types,
        doc=doc,
        funcs=funcs,
        funcs_map=funcs_map,
        attrs=attrs_obj,
        vars=get_vars(info, attrs_obj, doc[0]),
        lib_name=lib_name,
    )
    return obj


@dataclass
class TypeLib:
    name: str
    typeinfo: wint.PyITypeLib
    long_name: str
    attr: wint.TLIBATTR
    objects: dict[str, ComObject]
    ref_types: dict[str, wint.PyITypeInfo]

    # import_map: dict[str,list[str]]
    # import_libs: dict[str,wint.PyITypeLib]

    def __post_init__(self):
        self.import_map: dict[str, list[str]] = {}
        self.lib_map: dict[str, wint.PyITypeLib] = {}

        imports = {k: v for k, v in self.ref_types.items() if k not in self.objects}

        for k, v in imports.items():
            lib = v.GetContainingTypeLib()[0]
            lib_name = lib.GetDocumentation(-1)[0]
            if lib_name in self.import_map:
                self.import_map[lib_name].append(k)
            else:
                self.import_map[lib_name] = [k]
            if lib_name in self.lib_map:
                continue
            print(lib_name)
            self.lib_map[lib_name] = lib

    def pyi_str(self, py=False):
        txt = ""

        import_b = ""

        if not py:
            for k, v in self.import_map.items():
                if k == self.name:
                    for ref in v:
                        print(f"{k} problem with {ref}")
                import_b += f"import wincom_typegen.l_{k.lower()} as l_{k.lower()}\n"
                # for ref in v:
                #     txt += f'    {ref},\n'
                # txt += ')\n\n'
            txt += "\n\n"

        objs = list(self.objects.values())

        def sort_key(obj: ComObject):
            if obj.attrs.typekind == 6:
                return (0, obj.name)
            if obj.attrs.typekind in [1, 0]:
                return (1, obj.name)
            if obj.name[0] == "_":
                return (2, obj.name)
            if len(obj.impl_types) > 0:
                return (9, obj.name)
            else:
                return (3, obj.name)

        objs.sort(key=lambda x: sort_key(x))

        objs_listed = set()

        obj_dict = {i: obj for i, obj in enumerate(objs) if obj.name not in objs_listed}

        while len(obj_dict) > 0:
            # i = 0

            obj_dict = {
                i: obj for i, obj in enumerate(objs) if obj.name not in objs_listed
            }

            for i in obj_dict:
                if len(objs[i].impl_types) == 0:
                    obj: ComObject = obj_dict[i]
                    txt += obj.pyi_str(self.ref_types, list(self.objects.keys()), py)
                    objs_listed.add(obj.name)
                    continue
                else:
                    if all(x.name in objs_listed for x in objs[i].impl_types):
                        obj: ComObject = obj_dict[i]
                        txt += obj.pyi_str(
                            self.ref_types, list(self.objects.keys()), py
                        )
                        objs_listed.add(obj.name)
                        continue

            # txt += obj.pyi_str(self.ref_types,list(self.objects.keys()))

        import_txt = ""
        if py:
            import_txt += f"from wincom_typegen.typegen_classes import ComDummy\n"
            import_txt += f"import win32com.client as win\n"
        if "(Enum)" in txt:
            import_txt += f"from enum import Enum\n"
        if "datetime.date" in txt:
            import_txt += "import datetime\n"
        if "comtypes" in txt:
            import_txt += "import comtypes\n"
        if "ctypes" in txt:
            import_txt += "import ctypes\n"
        if "ctypes.wintypes" in txt:
            import_txt += "import ctypes.wintypes\n"

        import_txt += import_b
        import_txt += "\n\n"
        import_txt += "VbObject = object()\n"
        import_txt += "VbUnknown = object()\n"
        import_txt += "VbVariant = object()\n"

        if py:
            import_txt += "\n\n"
            for item in target_objects:
                if item[0] == self.name:
                    import_txt += f'def get_{item[0].lower()}_{item[1].lower()}() -> "{item[1]}":\n'
                    import_txt += f'    return win.Dispatch("{item[0]}.{item[1]}")\n\n'
        else:
            import_txt += "\n\n"
            for item in target_objects:
                if item[0] == self.name:
                    import_txt += f'def get_{item[0].lower()}_{item[1].lower()}() -> "{item[1]}": ...\n'

        desc = f"#{self.name}\n"
        desc += f"#{self.long_name}\n"
        desc += f"#{self.attr[0]}\n"

        return desc + import_txt + txt


class LibCollection:
    def __init__(self) -> None:
        self.lib_map: dict[str, TypeLib] = {}
        self.lib_map2: dict[str, (str, str)] = {}
        self.libs_to_get = [
            "Excel",
            "Access",
            "Scripting",
            "Outlook",
            "PowerPoint",
            "ADODB",
            "Shell32",
            "ADOX",
            "ADOR",
            "ADOMD",
            "Publisher",
            "Word",
            "VBA",
        ]
        pass

    def scan_for_type_libs(self):
        self.type_libs_on_system: list[wint.PyITypeLib] = []
        type_lib_specs = EnumTlbs(0)
        for tl in type_lib_specs:
            try:
                self.type_libs_on_system.append(LoadTypeLib(tl.dll))
            except:
                print(f"{tl.desc} failed to load")
                continue

        tl_docs = [
            [z for z in x.GetDocumentation(-1)]
            + [f"{x.GetLibAttr()[-3]}.{x.GetLibAttr()[-2]}", x]
            for x in self.type_libs_on_system
        ]
        tl_docs.sort(key=lambda x: (x[0], x[-2]), reverse=True)
        max_len = max([len(x[0]) for x in tl_docs]) + 2
        lib_names = [f"{x[0]: >{max_len}}: {x[1]} {x[-2]} {x[-1]}" for x in tl_docs]
        Path(__file__).parent.joinpath("libs_on_system.txt").write_text(
            "\n".join(lib_names)
        )

        self.lib_map_unprocessed: dict[str, wint.PyITypeLib] = {}
        for x in tl_docs:
            if x[0] in self.lib_map:
                continue
            self.lib_map_unprocessed[x[0]] = x[-1]

    def process_lib(self, lib: wint.PyITypeLib):
        if isinstance(lib, str):
            if not Path(lib).exists():
                print(f"{lib} not found")
                return
            try:
                lib = LoadTypeLib(lib)
            except:
                print(f"{lib} failed to load")
                return

        ref_types: dict[str, wint.PyITypeInfo] = {}

        lib_atter = lib.GetLibAttr()

        lib_doc = lib.GetDocumentation(-1)
        self.lib_map2[str(lib_atter[0])] = (lib_doc[0], lib_doc[1])
        lib_name = lib_doc[0]
        if lib_name in self.lib_map:
            # print(f'{lib_name} already processed')
            return
        obj_by_name: dict[str, ComObject] = {}

        for i in range(lib.GetTypeInfoCount()):
            obj = get_obj(i, lib, lib_name, ref_types)
            if not obj:
                continue
            # obj_by_id[i] = obj
            obj_by_name[obj.name] = obj

        tlib = TypeLib(
            name=lib_name,
            long_name=lib_doc[1],
            typeinfo=lib,
            attr=lib_atter,
            objects=obj_by_name,
            ref_types=ref_types,
        )

        self.lib_map[tlib.name] = tlib

        for ref in tlib.import_map:
            # if ref in self.lib_map:
            # continue
            print(ref)
            # ref_types.clear()
            self.process_lib(tlib.lib_map[ref])

    def write_lib_names_as_txt(self, path: Path):
        output_folder = path.joinpath("lib_names")
        output_folder.mkdir(exist_ok=True)
        for k, v in self.lib_map.items():
            txt = ""
            for i in range(v.typeinfo.GetTypeInfoCount()):
                # interafaces = []
                # for ii in range (v.typeinfo.GetTypeInfo(i).)
                txt += f"{i: >3}:  {v.typeinfo.GetDocumentation(i)[0]}    \n"
            file = output_folder.joinpath(f"{k}.txt")
            file.write_text(txt)

    def write_libs(self, path: Path = Path(__file__).parent):
        for lib in self.lib_map.values():
            path.joinpath(f"l_{lib.name.lower()}.pyi").write_text(lib.pyi_str())
            path.joinpath(f"l_{lib.name.lower()}.py").write_text(lib.pyi_str(True))
            # path.joinpath("__init__.py").write_text("")
        self.write_lib_names_as_txt(path)

    def process_lib_map(self):
        for tl in self.lib_map_unprocessed.values():
            try:
                if tl.GetDocumentation(-1)[0] in self.libs_to_get:
                    self.process_lib(tl)
                # col.process_lib(tl)
            except:
                print(f"{tl} failed to process")
                continue

    def get_running_com_objects(self):

        context = CreateBindCtx(0)
        running_coms = GetRunningObjectTable()
        objs:dict[str,wint.PyITypeLib] = {}
        print('running com objects:')
        for m in running_coms.EnumRunning():
            objid = m.GetDisplayName(context,m)
            obj = win.Dispatch(objid.strip('!'))
            lib = get_type_lib(obj)
            obj_name = lib.GetDocumentation(-1)[0]
            print('\t',obj_name)
            objs[obj_name] = lib
            
        return objs
# %%
