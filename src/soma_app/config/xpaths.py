CONFIRM_YES_XPATHS = (
    "//*[self::button or self::a][contains(normalize-space(.), 'Sim')]",
    "//*[self::button or self::a][contains(normalize-space(.), 'SIM')]",
    "//*[self::button or self::a][contains(normalize-space(.), 'Yes')]",
    "//*[self::button or self::a][contains(normalize-space(.), 'OK')]",
)

CONFIRM_NO_XPATHS = (
    "//*[self::button or self::a][contains(normalize-space(.), 'Não')]",
    "//*[self::button or self::a][contains(normalize-space(.), 'Nao')]",
    "//*[self::button or self::a][contains(normalize-space(.), 'No')]",
)
