"""Correspondencias explícitas Kobo -> HXL de las plantillas 345w."""

SECTORES = ('Proteccion', 'WASH', 'Educacion')
LIDER = 'UNICEF - Fondo de las Naciones Unidas para la Infancia'

VULNERABILIDAD = {
    '#reached+indigenous': 'vuln_indigena',
    '#reached+disabled': 'vuln_discapacidad',
    '#reached+f+pregnant': 'vuln_embarazadas',
    '#reached+lgbti': 'vuln_lgbti',
    '#victims+trafficking': 'vuln_victima_trata',
    '#reached+afro': 'vuln_afrodescendiente',
}

DESAGREGACIONES = {s: {} for s in (*SECTORES, 'Transversal')}
for sexo, infancia, adulto in (('f', 'ninas', 'mujeres'), ('m', 'ninos', 'hombres')):
    for edad in ('0_5', '6_9', '10_11', '12_14', '15_17'):
        DESAGREGACIONES['Proteccion'][f'#reached+{sexo}+children+age{edad}'] = [f'{infancia}_{edad}_prot']
    for edad in ('18_19', '20_49', '50_59'):
        DESAGREGACIONES['Proteccion'][f'#reached+{sexo}+adult+age{edad}'] = [f'{adulto}_{edad}_prot']
    DESAGREGACIONES['Proteccion'][f'#reached+{sexo}+elderly+age60'] = [f'{adulto}_60_mas_prot']
    for edad in ('0_2', '3_5', '6_11', '12_17'):
        DESAGREGACIONES['Educacion'][f'#reached+{sexo}+children+age{edad}'] = [f'{infancia}_{edad}']
    # Equivalencia solicitada por el usuario: 18-59 -> 20-59.
    DESAGREGACIONES['Educacion'][f'#reached+{sexo}+teachers+age20_59'] = [f'{adulto}_18_59']
    DESAGREGACIONES['Educacion'][f'#reached+{sexo}+adult+age20_59'] = [f'personal_comunitario_{adulto}_18_59']
    DESAGREGACIONES['Educacion'][f'#reached+{sexo}+teachers+age60_'] = [f'{adulto}_60_mas']
    DESAGREGACIONES['Educacion'][f'#reached+{sexo}+elderly+age60_'] = [f'personal_comunitario_{adulto}_60_mas']
    DESAGREGACIONES['Transversal'][f'#reached+{sexo}+children+age0_17'] = [f'{infancia}_0_17_transv']
    DESAGREGACIONES['Transversal'][f'#reached+{sexo}+adult+age18_59'] = [f'{adulto}_18_59_transv']
    DESAGREGACIONES['Transversal'][f'#reached+{sexo}+elderly+age60_'] = [f'{adulto}_60_mas_transv']
    DESAGREGACIONES['WASH'][f'#reached+{sexo}+children+age0_17'] = [f'{infancia}_0_17']
    DESAGREGACIONES['WASH'][f'#reached+{sexo}+adult+age18_59'] = [f'{adulto}_18_59']
    DESAGREGACIONES['WASH'][f'#reached+{sexo}+elderly+age60_'] = [f'{adulto}_60_79']
    DESAGREGACIONES['WASH'][f'#reached+{sexo}+elderly+age80_'] = [f'{adulto}_80mas']
for edad in ('0_17', '18_59', '60_'):
    DESAGREGACIONES['Transversal'][f'#reached+nobinario+age{edad}'] = [f'nobin_{"60_mas" if edad == "60_" else edad}_transv']
DESAGREGACIONES['Transversal']['#reached+noidentificado'] = ['no_identifica_transv']

# Sólo actividades transversales con correspondencia demostrable en los Excel.
# Las unidades de ACT_23 y ACT_75 no se pueden obtener del total de personas.
TRANSVERSAL = {
    'ACT_1': ('GCLCO/SCA13E', 'SCA13E/CIN1'),
    'ACT_22': ('GCLCO/CCA1E', 'CCA1E/CIN1'),
    'ACT_23': ('GCLCO/CCA3E', 'CCA3E/CIN3'),
    'ACT_75': ('GCLCO/SCA13E', 'SCA13E/CIN3'),
    'GEN_1': ('GCLCO/CCA4E', 'CCA4E/CIN1'),
    'GEN_2': ('GCLCO/CCA1E', 'CCA1E/CIN1'),
    'SBC_04': ('SBC.02', 'SBC.02'),
}

TIPOS_LUGAR = {
    'Proteccion': {
        'comunidad': 'Comunidad', 'escuela': 'Escuela o centro educativo',
        'espacio_amigable': 'Espacio o centro comunitario / institucional',
        'albergue': 'EAT - Espacio de alojamiento temporal (civil)',
        'albergue_onu': 'EAT - Espacio de alojamiento temporal (civil)',
    },
    'WASH': {
        'comunidad': 'Comunidad', 'escuela': 'Escuela o centro educativo',
        'albergue': 'EAT - Espacio de alojamiento temporal (civil)',
        'albergue_onu': 'EAT - Espacio de alojamiento temporal (civil)',
    },
    'Educacion': {
        'comunidad': 'Espacio o centro comunitario / institucional',
        'escuela': 'Escuela o centro educativo',
        'espacio_amigable': 'Espacio o centro comunitario / institucional',
        'albergue': 'EAT - Espacio de alojamiento temporal (civil)',
        'albergue_onu': 'EAT - Espacio de alojamiento temporal (civil)',
    },
}
