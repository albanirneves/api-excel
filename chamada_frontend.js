(async () => {
    //const API_URL = 'https://excel-creator.vercel.app/api/plan';
    const API_URL = 'http://localhost:3000/api/plan';
    
    const jsonRequest = {
        "font": { "name": "Calibri", "size": 11 },
        "columns": [
            { "width": 11 },
            { "width": 30 },
            { "width": 13 },
            { "width": 11 },
            { "width": 11 },
            { "width": 11 },
            { "width": 11 },
            { "width": 11 },
            { "width": 11 },
            { "width": 11 },
            { "width": 15 },
            { "width": 15 },
            { "width": 15 },
            { "width": 18 },
            { "width":  8 },
            { "width": 18 },
            { "width": 18 },
            { "width": 18 },
            { "width": 18 },
            { "width": 10 },
            { "width": 10 },
            { "width": 10 },
            { "width": 10 },
            { "width": 50 }
        ],
        "rows": [
            {
                "height": 33,
                "font": { "name": "Calibri", "size": 14, "bold": true },
                "alignment": { "vertical": "middle", "horizontal": "center" },
                "cells": [
                    {
                        "ref": "A1:X1",
                        "value": "VISITA DE INSPEÇÃO TÉCNICA DO PAVIMENTO"
                    }
                ]
            },
            {
                "height": 15,
                "font": { "bold": true },
                "alignment": { "vertical": "middle", "horizontal": "center" },
                "cells": [
                    {
                        "ref": "A2:G2",
                        "value": "INFORMAÇÕES DO TRECHO"
                    },
                    {
                        "ref": "H2:J2",
                        "value": "POSIÇÃO FOTO"
                    },
                    {
                        "ref": "K2:O2",
                        "value": "INFORMAÇÕES PISTA"
                    },
                    {
                        "ref": "P2:S2",
                        "value": "COMPOSIÇÃO DO PAVIMENTO"
                    },
                    {
                        "ref": "T2:W3",
                        "value": "DEFEITOS NO PAVIMENTO"
                    },
                    {
                        "ref": "X2:X3",
                        "value": "LINK FOTO"
                    }
                ]
            },
            {
                "height": 15,
                "font": { "bold": true },
                "alignment": { "vertical": "middle", "horizontal": "center" },
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
                    },
                    {
                        "ref": "F3",
                        "value": "Pista"
                    },
                    {
                        "ref": "G3",
                        "value": "Faixa"
                    },
                    {
                        "ref": "H3",
                        "value": "Nº Foto"
                    },
                    {
                        "ref": "I3",
                        "value": "Km Foto"
                    },
                    {
                        "ref": "J3",
                        "value": "Estaca Foto"
                    },
                    {
                        "ref": "K3",
                        "value": "Largura Faixa 1"
                    },
                    {
                        "ref": "L3",
                        "value": "Largura Faixa 2"
                    },
                    {
                        "ref": "M3",
                        "value": "Acostamento"
                    },
                    {
                        "ref": "N3",
                        "value": "Larg. Acostamento"
                    },
                    {
                        "ref": "O3",
                        "value": "Seção"
                    },
                    {
                        "ref": "P3",
                        "value": "Revestimento"
                    },
                    {
                        "ref": "Q3",
                        "value": "Base"
                    },
                    {
                        "ref": "R3",
                        "value": "Sub-base"
                    },
                    {
                        "ref": "S3",
                        "value": "Subleito"
                    }
                ]
            },
            {
                "height": 15,
                "alignment": { "vertical": "middle", "horizontal": "left" },
                "cells": [
                    {
                        "ref": "A4",
                        "value": "120EBA0001"
                    },
                    {
                        "ref": "B4",
                        "value": "Santa Terezinha - Elísio Medrado"
                    },
                    {
                        "ref": "C4",
                        "alignment": { "horizontal": "right" },
                        "value": 0
                    },
                    {
                        "ref": "D4",
                        "alignment": { "horizontal": "right" },
                        "value": 0
                    },
                    {
                        "ref": "E4",
                        "value": "Direito"
                    },
                    {
                        "ref": "F4",
                        "value": "Dupla"
                    },
                    {
                        "ref": "G4",
                        "alignment": { "horizontal": "right" },
                        "value": 1
                    },
                    {
                        "ref": "H4",
                        "alignment": { "horizontal": "right" },
                        "value": 1
                    },
                    {
                        "ref": "I4",
                        "alignment": { "horizontal": "right" },
                        "value": 0.523
                    },
                    {
                        "ref": "J4",
                        "value": "26+03"
                    },
                    {
                        "ref": "K4",
                        "alignment": { "horizontal": "right" },
                        "value": 3.6
                    },
                    {
                        "ref": "L4",
                        "alignment": { "horizontal": "right" },
                        "value": 3.6
                    },
                    {
                        "ref": "M4",
                        "value": "Presente"
                    },
                    {
                        "ref": "N4",
                        "alignment": { "horizontal": "right" },
                        "value": 2.5
                    },
                    {
                        "ref": "O4",
                        "value": "Aterro"
                    },
                    {
                        "ref": "P4",
                        "value": "CBUQ"
                    },
                    {
                        "ref": "Q4",
                        "value": "Cascalho Laterítico"
                    },
                    {
                        "ref": "R4",
                        "value": "Cascalho Quartzoso"
                    },
                    {
                        "ref": "S4",
                        "value": "Cascalho Quartzoso"
                    },
                    {
                        "ref": "T4",
                        "value": "ATC"
                    },
                    {
                        "ref": "U4",
                        "value": "Exsudação"
                    },
                    {
                        "ref": "V4",
                        "value": "Panela"
                    },
                    {
                        "ref": "W4",
                        "value": ""
                    },
                    {
                        "ref": "X4",
                        "value": "http://..."
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