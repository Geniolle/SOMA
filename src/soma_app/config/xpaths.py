CONFIRM_POPUP_CONTAINER_XPATHS = (
    "//div[contains(@class,'swal2-popup') and (contains(normalize-space(.), 'realizar o pagamento') or contains(normalize-space(.), 'deseja realizar o pagamento'))]",
    "//div[contains(@class,'swal2-container') and (contains(normalize-space(.), 'realizar o pagamento') or contains(normalize-space(.), 'deseja realizar o pagamento'))]",
    "//li[.//*[contains(normalize-space(.), 'realizar o pagamento') or contains(normalize-space(.), 'deseja realizar o pagamento')]]",
)

CONFIRM_YES_XPATHS = (
    "//div[contains(@class,'swal2-container')]//button[normalize-space()='Sim']",
    "//div[contains(@class,'swal2-container')]//a[normalize-space()='Sim']",
    "//li[.//*[contains(normalize-space(.), 'realizar o pagamento') or contains(normalize-space(.), 'deseja realizar o pagamento')]]//button[normalize-space()='Sim']",
    "//li[.//*[contains(normalize-space(.), 'realizar o pagamento') or contains(normalize-space(.), 'deseja realizar o pagamento')]]//a[normalize-space()='Sim']",
    "//*[self::button or self::a][contains(normalize-space(.), 'Sim')]",
    "//*[self::button or self::a][contains(normalize-space(.), 'SIM')]",
    "//*[self::button or self::a][contains(normalize-space(.), 'Yes')]",
    "//*[self::button or self::a][contains(normalize-space(.), 'Yes!')]",
    "//*[self::button or self::a][contains(normalize-space(.), 'OK')]",
    "//*[self::button or self::a][contains(normalize-space(.), 'Ok')]",
    "//*[self::button or self::a][contains(normalize-space(.), 'OK!')]",
    "//*[self::button or self::a][contains(normalize-space(.), 'Ok!')]",
    "//*[self::button or self::a][contains(normalize-space(.), 'Confirmar')]",
    "//*[self::button or self::a][contains(normalize-space(.), 'Aceitar')]",
    "/html/body/ul/li/div/div[2]/button[1]",
)

CONFIRM_NO_XPATHS = (
    "//*[self::button or self::a][contains(normalize-space(.), 'Não')]",
    "//*[self::button or self::a][contains(normalize-space(.), 'Nao')]",
    "//*[self::button or self::a][contains(normalize-space(.), 'No')]",
)
