(async () => {
    const API_URL = 'https://excel-creator.vercel.app/api/plan';
    
    const jsonRequest = {
        //aqui vc pode colocar as celular que quer que aparece filtro
        "autoFilter": {
            "from": "A3",
            "to": "C3"
        },

        //aqui é a fonte que se aplica a toda a planilha
        //mas ela é sobrescrita se declarar novamente nas linhas ou célular
        "font": { 
            "name": "Calibri", 
            "size": 11 
        },

        //largura das colunas
        "columns": [
            { "width": 20 },
            { "width": 30 },
            { "width": 18 },
            { "width": 20 },
            { "width": 20 }
        ],

        //lista de linhas
        "rows": [
            {
                "height": 33,//altura da linha
                "hidden": false,//true a coluna fica escondida
                "font": { 
                    "name": "Calibri", 
                    "size": 14, 
                    "bold": true, //negrito
                    "underline": true, //sublinhado
                    "italic": true, //italico

                    //aqui vc manda a cor em hexadecimal
                    //para facilitar visite: http://www.color-hex.com/
                    "color": { "argb": "FF00FF" }  
                },

                //alinhamento da linha, também pode ser declarado nas célular individuais
                "alignment": { 
                    "vertical": "middle", 
                    "horizontal": "center" 
                },

                //lista de células da linha
                "cells": [
                    {
                        //se colocar um intervalo, as células serão mescladas
                        "ref": "A1:E1",
                        "value": "VISITA DE INSPEÇÃO",

                        //opcao de borda comum
                        "border": {
                            "top"   : { "style": "thin" },
                            "left"  : { "style": "thin" },
                            "bottom": { "style": "thin" },
                            "right" : { "style": "thin" }
                        }
                    }
                ]
            },
            {
                "height": 15,
                "font": { 
                    "bold": true,
                    "strike": true//tachado
                },
                "alignment": { 
                    "vertical": "middle", 
                    "horizontal": "center" 
                },
                "cells": [
                    {
                        "ref": "A2:C2",
                        "value": "POSIÇÃO FOTO"
                    },
                    {
                        "ref": "D2:E2",
                        "value": "DEFEITOS NO PAVIMENTO"
                    }
                ]
            },
            {
                "height": 15,
                "font": { 
                    "bold": true 
                },
                "alignment": { 
                    "vertical": "middle", 
                    "horizontal": "center" 
                },
                "cells": [
                    {
                        "ref": "A3",
                        "value": "SNV"
                    },
                    {
                        "ref": "B3",
                        "value": "Trecho"
                    },
                    {
                        "ref": "C3",
                        "value": "Estaca Inicial"
                    },
                    {
                        "ref": "D3",
                        "value": "Km inicial"
                    },
                    {
                        "ref": "E3",
                        "value": "Lado"
                    }
                ]
            },
            {
                "height": 15,
                "alignment": { 
                    "vertical": "middle", 
                    "horizontal": "left" 
                },
                "cells": [
                    {
                        "ref": "A4",
                        "alignment": { 
                            "horizontal": "right" 
                        },
                        "value": "01/01/2021"
                    },
                    {
                        "ref": "B4",
                        "value": "Santa Terezinha - Elísio Medrado"
                    },
                    {
                        "ref": "C4",
                        "alignment": { 
                            "horizontal": "right" 
                        },
                        "value": 0//valores numericos serão tratados como numeros apenas sem aspas
                    },
                    {
                        "ref": "D4",
                        "alignment": { 
                            "horizontal": "right" 
                        },
                        "value": 0
                    },
                    {
                        "ref": "E4",
                        "value": "Direito"
                    }
                ]
            },
            {
                "height": 15,
                "alignment": { 
                    "vertical": "middle", 
                    "horizontal": "left" 
                },
                "cells": [
                    {
                        "ref": "A5",
                        "alignment": { 
                            "horizontal": "right" 
                        },
                        "value": "Teste"
                    },
                    {
                        "ref": "B5",
                        "value": "Santa Terezinha - Elísio Medrado"
                    },
                    {
                        "ref": "C5",
                        "alignment": { 
                            "horizontal": "right" 
                        },
                        "value": 0
                    },
                    {
                        "ref": "D5",
                        "alignment": { 
                            "horizontal": "right" 
                        },
                        "value": 0
                    },
                    {
                        "ref": "E5",
                        "value": "Direito"
                    }
                ]
            },
            {
                "height": 15,
                "alignment": { 
                    "vertical": "middle", 
                    "horizontal": "left" 
                },
                "cells": [
                    {
                        "ref": "A7",
                        "value": "Valor"
                    },
                    {
                        "ref": "C7",
                        "value": "bordas",
                        
                        //bordas grossas
                        "border": {
                            "top"   : { "style": "medium" },
                            "left"  : { "style": "medium" },
                            "bottom": { "style": "medium" },
                            "right" : { "style": "medium" }
                        }
                    },
                    {
                        "ref": "E7",
                        "value": "bordas",

                        //bordas coloridas
                        "border": {
                            "top"   : {"style": "thin", "color": { "argb": "FF00FF" } },
                            "left"  : {"style": "thin", "color": { "argb": "FF00FF" } },
                            "bottom": {"style": "thin", "color": { "argb": "FF00FF" } },
                            "right" : {"style": "thin", "color": { "argb": "FF00FF" } }
                        }
                    },
                    {
                        "ref": "G7",
                        "value": "bordas",

                        //bordas duplas
                        "border": {
                            "top"   : {"style": "double", "color": { "argb": "00FFFF" } },
                            "left"  : {"style": "double", "color": { "argb": "00FFFF" } },
                            "bottom": {"style": "double", "color": { "argb": "00FFFF" } },
                            "right" : {"style": "double", "color": { "argb": "00FFFF" } }
                        }
                    },
                    {
                        "ref": "I7",
                        "value": "bordas",

                        //bordas pontilhadas
                        "border": {
                            "top"   : {"style": "mediumDashed" },
                            "left"  : {"style": "mediumDashed" },
                            "bottom": {"style": "mediumDashed" },
                            "right" : {"style": "mediumDashed" }
                        }
                    }
                ]
            },
            {
                "height": 15,
                "alignment": { 
                    "vertical": "middle", 
                    "horizontal": "left" 
                },
                "cells": [
                    {
                        "ref": "A9",
                        "value": "Valor"
                    },
                    {
                        "ref": "C9",
                        "value": "preencher",

                        //cor da fonte
                        "font": {
                            "color": { "argb": "FF00FF" }
                        },

                        //cor de fundo, só colocar a cor hexadecimal
                        "fill": {
                            "type"   : "pattern",
                            "pattern": "solid",
                            "bgColor": { "argb": "8B0000" }
                        }
                    },
                    {
                        "ref": "E9",
                        //imagem vc tem duas opcoes, aqui vc pode mandar o link da imagem ou
                        //pode tambem converter a imagem pra base64 e mandar o texto aqui
                        //no lugar do link
                        "image": {
                            "link": "https://img.ibxk.com.br/2020/01/30/30021141299110.jpg",
                            "width": 250, 
                            "height": 100
                        }
                    }
                ]
            }
        ]
    };
    
    try {
        const myHeaders = new Headers();
        
        myHeaders.append('Content-Type', 'application/json');
        
        //faz a chamada à API (a variavel jsonRequest contem o json com a estrutura)
        const res = await fetch(API_URL, { 
            method: 'POST',
            headers: myHeaders,
            body: JSON.stringify(jsonRequest)
        })
    
        if (!res.ok) {
            throw new Error(res.statusText);
        } else {
            //primeiro converte a resposta de base 64 para um blob
            const bin = atob(await res.text());
            const arrayBuffer = new ArrayBuffer(bin.length);
            const view = new Uint8Array(arrayBuffer);
            for (let i = 0; i != bin.length; ++i) view[i] = bin.charCodeAt(i) & 0xFF;
            const blob = new Blob([arrayBuffer], { 
                type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;' 
            });
    
            //depois faz o download
            const link = window.document.createElement('a');
            link.href = window.URL.createObjectURL(blob);
            link.download = `demo.xlsx`;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    
    } catch(error) {
        throw new Error('Error occurred:' + error);
    }
})();