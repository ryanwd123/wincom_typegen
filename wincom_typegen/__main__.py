#%%
from .gen import LibCollection
#%%

if __name__ == "__main__":
    col = LibCollection()
    col.scan_for_type_libs()
    col.process_lib_map()
    col.write_libs()