module.exports = {
  headerJSONFormatter: {
    elmType: 'div',
    attributes: {
      class: 'ms-borderColor-neutralTertiary'
    },
    style: {
      width: '99%',
      'border-top-width': '0px',
      'border-bottom-width': '1px',
      'border-left-width': '0px',
      'border-right-width': '0px',
      'border-style': 'solid',
      'margin-bottom': '16px'
    },
    children: [
      {
        elmType: 'div',
        style: {
          display: 'flex',
          'box-sizing': 'border-box',
          'align-items': 'center'
        },
        children: [
          {
            elmType: 'div',
            attributes: {
              iconName: 'Calendar',
              class: 'ms-fontSize-42 ms-fontWeight-regular ms-fontColor-themePrimary',
              title: 'Details'
            },
            style: {
              flex: 'none',
              padding: '0px',
              'padding-left': '0px',
              height: '36px'
            }
          }
        ]
      },
      {
        elmType: 'div',
        attributes: {
          class: 'ms-fontColor-neutralSecondary ms-fontWeight-bold ms-fontSize-24'
        },
        style: {
          'box-sizing': 'border-box',
          width: '100%',
          'text-align': 'left',
          padding: '21px 12px',
          overflow: 'hidden'
        },
        children: [
          {
            elmType: 'div',
            txtContent: "='Saksinformasjon: ' + [$Title]"
          }
        ]
      }
    ]
  },
  footerJSONFormatter: '',
  bodyJSONFormatter: {
    sections: [
      {
        displayname: '',
        fields: [
          'Møtedato',
          'Sakstype',
          'Status',
          'Oppmeldt av',
          'Sortering',
          'Tidsbruk'
        ]
      },
      {
        displayname: 'Tittel og beskrivelse',
        fields: [
          'Sakstittel',
          'Beskrivelse av sak',
          'Sakstittel'
        ]
      },
      {
        displayname: 'Behandling',
        fields: [
          'Kommentarer',
          'Beslutning',
          'Ansvarlig for oppfølging'
        ]
      },
      {
        displayname: 'Publisering - ikke aktivert',
        fields: [
          'Publisere referat',
          'Publisere vedlegg'
        ]
      },
      {
        displayname: 'Arkivinformasjon - ikke aktivert',
        fields: [
          'Dokumentnummer',
          'Arkiveringsstatus',
          'Arkiver på nytt'
        ]
      },
      {
        displayname: 'Annet',
        fields: [
          'smart_UnntattOffentlighet',
          'smart_DokumentNummer',
          'smart_Arkiveringsstatus',
          'smart_Vedlegglenke',
          'smart_ReferatID',
          'smart_Elementversjon'
        ]
      }
    ]
  }
}
