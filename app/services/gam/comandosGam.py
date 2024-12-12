class CorreosDocentesAdministrativos:
    def getCreateGroup(self, group):
        return f"gam create group {group}"
    
    def fillGroupWithCsv(self, csv):
        return f"gam csv {csv} gam update group \"~Group Email\" add \"~Member Role\" \"~Member Email\""
        