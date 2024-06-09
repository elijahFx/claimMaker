let selectedOption = ""
const { Document, Packer, Paragraph, TextRun, AlignmentType, Table, TableRow, TableCell, WidthType, BorderStyle } = docx;
const INDENT = 4952

document.addEventListener("DOMContentLoaded", function() {

    let vozvr = true
    let zamena = false
    let sorazm = false
    let vozmesh = false
    let ustranen = false
    let daysForOtvet = 7

    const sorazmAmountDiv = document.querySelector("#sorazmDiv")
    const vozmAmountDiv = document.querySelector("#vozmeshDiv")
    const options = document.getElementById('options')


    options.addEventListener('change', function() {
        selectedOption = this.options[this.selectedIndex].text;
        console.log("Selected option: " + selectedOption);

        switch (selectedOption) {
            case "Возврат денежных средств":
                vozvr = true
                zamena = false
                sorazm = false
                vozmesh = false
                ustranen = false
                vozmAmountDiv.style.display = `none`
                sorazmAmountDiv.style.display = `none`
                break;
                case "Замена товара":
                vozvr = false
                zamena = true
                sorazm = false
                vozmesh = false
                ustranen = false
                vozmAmountDiv.style.display = `none`
                sorazmAmountDiv.style.display = `none`
                break;
                case "Безвозмездное устранение недостатков товара":
                vozvr = false
                zamena = false
                sorazm = false
                vozmesh = false
                ustranen = true
                vozmAmountDiv.style.display = `none`
                sorazmAmountDiv.style.display = `none`
                break;
                case "Соразмерное уменьшение стоимости товара":
                    console.log("hello");
                vozvr = false
                zamena = false
                sorazm = true
                vozmesh = false
                ustranen = false
                vozmAmountDiv.style.display = `none`
                sorazmAmountDiv.style.display = `block`
                break;
                case "Возмещение расходов на устранение недостатков товара третьими лицами":
                vozvr = false
                zamena = false
                sorazm = false
                vozmesh = true
                ustranen = false
                sorazmAmountDiv.style.display = `none`
                vozmAmountDiv.style.display = `block`
                break;
            default:
                break;
        }



        if(vozvr || sorazm || vozmesh) {
            daysForOtvet = 7
            console.log(daysForOtvet);
        } else if (ustranen || zamena) {
            daysForOtvet = 14
            console.log(daysForOtvet);
        }


    });

    


    document.querySelector("button").addEventListener("click", function() {
        const name = document.querySelector("#name").value;
        const address = document.querySelector("#address").value;
        const phone = document.querySelector("#phone").value;
        const liabelee = document.querySelector("#liabelee").value;
        const liabeleeAddress = document.querySelector("#liabeleeAddress").value;
        const complaint = document.querySelector("#complaint").value;
        const unp = document.querySelector("#unp").value;
        const price = document.querySelector("#price").value;
        const good = document.querySelector("#good").value;
        const firstDate = document.querySelector("#date").value;
        const sorazmAmount = document.querySelector("#sorazm").value
        const vozmAmount = document.querySelector("#vozmesh").value
        
        const abr = abbreviateName(name)
        const date = getCurrentDate()


        let paragraphs = []
        let paragraphs0 = []
        let boldTreb = []


        console.log(selectedOption);

            if(selectedOption === "Возврат денежных средств" || selectedOption === "") {
                paragraphs.push(new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    indent: { firstLine: 708 },
                    children: [
                        new TextRun({
                            text: `1. Расторгнуть заключенный, между нами, договор;`,
                            font: "Times New Roman",
                            size: 28, // 14 pt = 28 half-points
                        }),
                    ],
                }))
                paragraphs.push(
                new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    indent: { firstLine: 708 },
                    children: [
                        new TextRun({
                            text: `2. Вернуть уплаченные мной денежные средства по договору в размере ${price} белорусских рублей.`,
                            font: "Times New Roman",
                            size: 28, // 14 pt = 28 half-points
                        }),
                    ],
                    spacing: { after: 300 },
                })
            )
            console.log(paragraphs);
            
            } else if(selectedOption === "Замена товара") {
              
                paragraphs.push(new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    indent: { firstLine: 708 },
                    children: [
                        new TextRun({
                            text: `1. Осуществить замену товара ненадлежащего качества на товар надлежащего качества в срок не позднее 14 дней с момента получения настоящей претензии.`,
                            font: "Times New Roman",
                            size: 28, // 14 pt = 28 half-points
                        }),
                    ],
                    spacing: { after: 300 }
                })
            )
            boldTreb.push(new TextRun({
                alignment: AlignmentType.JUSTIFIED,
                indent: { firstLine: 708 },
                children: [
                    new TextRun({
                        text: `1.1. замены недоброкачественного товара товаром надлежащего качества;`,
                        font: "Times New Roman",
                        size: 28, // 14 pt = 28 half-points
                        bold: true
                    }),
                ],
            }))
            } else if(selectedOption === "Безвозмездное устранение недостатков товара") {
                
                paragraphs.push(new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    indent: { firstLine: 708 },
                    children: [
                        new TextRun({
                            text: `1. Безвозмездно устранить вышеизложенные недостатки товара не позднее 14 дней с момента получения настоящей претензии.`,
                            font: "Times New Roman",
                            size: 28, // 14 pt = 28 half-points
                        }),
                    ],
                    spacing: { after: 300 }
                })
            )
            boldTreb.push(new TextRun({
                alignment: AlignmentType.JUSTIFIED,
                indent: { firstLine: 708 },
                children: [
                    new TextRun({
                        text: `1.3. незамедлительного безвозмездного устранения недостатков товара;`,
                        font: "Times New Roman",
                        size: 28, // 14 pt = 28 half-points
                        bold: true
                    }),
                ],
            }))
            } else if(selectedOption === "Соразмерное уменьшение стоимости товара") {
               
                paragraphs.push(new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    indent: { firstLine: 708 },
                    children: [
                        new TextRun({
                            text: `1. Соразмерно уменьшить стоимость товара на ${sorazmAmount} белорусских рублей.`,
                            font: "Times New Roman",
                            size: 28, // 14 pt = 28 half-points
                        }),
                    ],
                    spacing: { after: 300 }
                })
            )
            paragraphs0.push(new Paragraph({
                alignment: AlignmentType.JUSTIFIED,
                indent: { firstLine: 708 },
                children: [
                    new TextRun({
                        text: `Поскольку мне был реализован некачественный товар, полагаю правомерным и необходимым соразмерно уменьшить его стоимость на ${sorazmAmount} белорусских рублей.`,
                        font: "Times New Roman",
                        size: 28, // 14 pt = 28 half-points
                    }),
                ],
            }))


            boldTreb.push(new TextRun({
                alignment: AlignmentType.JUSTIFIED,
                indent: { firstLine: 708 },
                children: [
                    new TextRun({
                        text: `1.2. соразмерного уменьшения покупной цены товара;`,
                        font: "Times New Roman",
                        size: 28, // 14 pt = 28 half-points
                        bold: true
                    }),
                ],
            }))


            } else if(selectedOption === "Возмещение расходов на устранение недостатков товара третьими лицами") {
            
                paragraphs.push(new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    indent: { firstLine: 708 },
                    children: [
                        new TextRun({
                            text: `1. Возместить расходы на устранение недостатка(-ов) товара общей стоимостью в ${vozmAmount} белорусских рублей.`,
                            font: "Times New Roman",
                            size: 28, // 14 pt = 28 half-points
                        }),
                    ],
                    spacing: { after: 300 }
                })
            )
            paragraphs0.push(new Paragraph({
                alignment: AlignmentType.JUSTIFIED,
                indent: { firstLine: 708 },
                children: [
                    new TextRun({
                        text: `В связи с наличием в товаре производственного(-ых) недостатка(-ов), мною были затрачены денежные средства в размере ${vozmAmount} белорусских рублей.`,
                        font: "Times New Roman",
                        size: 28, // 14 pt = 28 half-points
                    }),
                ],
            }))


            boldTreb.push(new TextRun({
                alignment: AlignmentType.JUSTIFIED,
                indent: { firstLine: 708 },
                children: [
                    new TextRun({
                        text: `1.4. возмещения расходов по устранению недостатков товара.`,
                        font: "Times New Roman",
                        size: 28, // 14 pt = 28 half-points
                        bold: true
                    }),
                ],
            }))


            }

        const consumer = {
            type: selectedOption,
            name,
            address,
            phone,
            complaint,
            date,
            good,
            price,
            firstDate,
            liabelee,
        }

/*

НА БУДУЩЕЕ, ЕСЛИ ЕГР ВДРУГ ЗАХОЧЕТ РАБОТАТЬ

        if (unp) {
            fetch(`https://egr.gov.by/api/v2/egr/getBaseInfoByRegNum/${unp}`)
                .then(response => response.json())
                .then(data => {
                    console.log(data);
                })
                .catch(error => {
                    console.error('Ошибка получения данных:', error);
                });
        }
*/


        const doc = new Document({
            sections: [
                {
                    properties: {
                        page: {
                            margin: {
                                    top: 1417, // 2 cm converted to inches
                                    right: 1063, // 1.5 cm converted to inches
                                    bottom: 1417, // 2 cm converted to inches
                                    left: 1575, // 3 cm converted to inches
                            }
                        }
                     },
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.LEFT,
                            children: [
                                new TextRun({
                                    text: `${liabelee}`,
                                    font: "Times New Roman",
                                    size: 28, // 14 pt = 28 half-points
                                }),
                            ],
                            indent: {left: INDENT}
                        }),
                        new Paragraph({
                            alignment: AlignmentType.LEFT,
                            children: [
                                new TextRun({
                                    text: `${liabeleeAddress}`,
                                    font: "Times New Roman",
                                    size: 28, // 14 pt = 28 half-points
                                }),
                            ],
                            indent: {left: INDENT}
                        }),
                        new Paragraph({
                            alignment: AlignmentType.LEFT,
                            children: [
                                new TextRun({
                                    text: `УНП - ${unp}`,
                                    font: "Times New Roman",
                                    size: 28, // 14 pt = 28 half-points
                                }),
                            ],
                            indent: {left: INDENT},
                            spacing: { after: 300 }
                        }),
                        new Paragraph({
                            alignment: AlignmentType.LEFT,
                            children: [
                                new TextRun({
                                    text: `${name}`,
                                    font: "Times New Roman",
                                    size: 28, // 14 pt = 28 half-points
                                }),
                            ],
                            indent: {left: INDENT}
                        }),
                        new Paragraph({
                            alignment: AlignmentType.LEFT,
                            children: [
                                new TextRun({
                                    text: `${address}`,
                                    font: "Times New Roman",
                                    size: 28, // 14 pt = 28 half-points
                                }),
                            ],
                            indent: {left: INDENT}
                        }),
                        new Paragraph({
                            alignment: AlignmentType.LEFT,
                            children: [
                                new TextRun({
                                    text: `${phone}`,
                                    font: "Times New Roman",
                                    size: 28, // 14 pt = 28 half-points
                                }),
                            ],
                            indent: {left: INDENT},
                            spacing: { after: 300 }
                        }),
                        new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                                new TextRun({
                                    text: `ПРЕТЕНЗИЯ`,
                                    bold: true,
                                    font: "Times New Roman",
                                    size: 28, // 14 pt = 28 half-points
                                }),
                            ],
                            spacing: { after: 300 }
                        }),
                        new Paragraph({
                            alignment: AlignmentType.JUSTIFIED,
                            indent: { firstLine: 708 },
                            children: [
                                new TextRun({
                                    text: `Я, ${name}, ${firstDate} заключил(-а) с Вашей организацией договор купли-продажи: ${good} (далее – товар).`,
                                    font: "Times New Roman",
                                    size: 28, // 14 pt = 28 half-points
                                }),
                            ],
                        }),
                        new Paragraph({
                            alignment: AlignmentType.JUSTIFIED,
                            indent: { firstLine: 708 },
                            children: [
                                new TextRun({
                                    text: `Я произвел(-а) оплату в сумме: ${price} белорусских рублей.`,
                                    font: "Times New Roman",
                                    size: 28, // 14 pt = 28 half-points
                                }),
                            ],
                        }),
                        new Paragraph({
                            alignment: AlignmentType.JUSTIFIED,
                            indent: { firstLine: 708 },
                            children: [
                                new TextRun({
                                    text: `В процессе эксплуатации товара в соответствии с его назначением и правилами пользования, мною был(-и) обнаружен(-ы) следующий(-ие) недостатки товара: ${complaint}.`,
                                    font: "Times New Roman",
                                    size: 28, // 14 pt = 28 half-points
                                }),
                            ],
                        }),
                        new Paragraph({
                            alignment: AlignmentType.JUSTIFIED,
                            indent: { firstLine: 708 },
                            children: [
                                new TextRun({
                                    text: `Таким образом, полагаю, что Вашей организацией было нарушено мое право, как потребителя, на товар надлежащего качества.`,
                                    font: "Times New Roman",
                                    size: 28, // 14 pt = 28 half-points
                                }),
                            ],
                        }),
                        ...paragraphs0,
                        new Paragraph({
                            alignment: AlignmentType.JUSTIFIED,
                            indent: { firstLine: 708 },
                            children: [
                                new TextRun({
                                    text: `Обращаю Ваше внимание, что согласно пункту 2, 4, 5 статьи 11 Закона «О защите прав потребителей» от 9 января 2002 г. № 90-З (далее – Закон), `,
                                    font: "Times New Roman",
                                    size: 28, // 14 pt = 28 half-points
                                }),
                            ],
                        }),
                        new Paragraph({
                            alignment: AlignmentType.JUSTIFIED,
                            indent: { firstLine: 708 },
                            children: [
                                new TextRun({
                                    text: `Продавец (исполнитель) обязан продемонстрировать работоспособность товара (результата работы) и передать потребителю товар (выполнить работу, оказать услугу), качество которого соответствует предоставленной информации о товаре (работе, услуге), требованиям законодательства, технических регламентов Таможенного союза, технических регламентов Евразийского экономического союза и условиям договора, а также по требованию потребителя предоставить ему необходимые средства измерений, прошедшие метрологический контроль в соответствии с законодательством об обеспечении единства измерений, документы, подтверждающие качество товара (результата работы, услуги), его комплектность, количество.`,
                                    font: "Times New Roman",
                                    size: 28, // 14 pt = 28 half-points
                                }),
                            ],
                        }),
                        new Paragraph({
                            alignment: AlignmentType.JUSTIFIED,
                            indent: { firstLine: 708 },
                            children: [
                                new TextRun({
                                    text: `Если продавец (исполнитель) при заключении договора был поставлен потребителем в известность о конкретных целях приобретения товара (выполнения работы, оказания услуги), продавец (исполнитель) обязан передать потребителю товар (выполнить работу, оказать услугу) надлежащего качества, пригодный для использования в соответствии с этими целями.`,
                                    font: "Times New Roman",
                                    size: 28, // 14 pt = 28 half-points
                                }),
                            ],
                        }),
                        new Paragraph({
                            alignment: AlignmentType.JUSTIFIED,
                            indent: { firstLine: 708 },
                            children: [
                                new TextRun({
                                    text: `При реализации потребителю товаров (выполнении работ, оказании услуг) по образцам, описаниям товаров (работ, услуг), содержащимся в каталогах, проспектах, рекламе, буклетах или представленным в фотографиях или иных информационных источниках, в том числе в глобальной компьютерной сети Интернет, продавец (исполнитель) обязан передать потребителю товар (выполнить работу, оказать услугу), качество которого соответствует таким образцам, описаниям, а также требованиям законодательства, технических регламентов Таможенного союза, технических регламентов Евразийского экономического союза.`,
                                    font: "Times New Roman",
                                    size: 28, // 14 pt = 28 half-points
                                }),
                            ],
                        }),
                        new Paragraph({
                            alignment: AlignmentType.JUSTIFIED,
                            indent: { firstLine: 708 },
                            children: [
                                new TextRun({
                                    text: `Также согласно пункту 1 статьи 20 Закона в случае реализации товара ненадлежащего качества, если его недостатки не были оговорены продавцом, потребитель вправе по своему выбору потребовать:`,
                                    font: "Times New Roman",
                                    size: 28, // 14 pt = 28 half-points
                                }),
                            ],
                        }),
                        new Paragraph({
                            alignment: AlignmentType.JUSTIFIED,
                            indent: { firstLine: 708 },
                            children: [
                                new TextRun({
                                    text: `1.1. замены недоброкачественного товара товаром надлежащего качества;`,
                                    font: "Times New Roman",
                                    size: 28, // 14 pt = 28 half-points
                                }),
                            ],
                        }),
                        new Paragraph({
                            alignment: AlignmentType.JUSTIFIED,
                            indent: { firstLine: 708 },
                            children: [
                                new TextRun({
                                    text: `1.2. соразмерного уменьшения покупной цены товара;`,
                                    font: "Times New Roman",
                                    size: 28, // 14 pt = 28 half-points
                                }),
                            ],
                        }),
                        new Paragraph({
                            alignment: AlignmentType.JUSTIFIED,
                            indent: { firstLine: 708 },
                            children: [
                                new TextRun({
                                    text: `1.3. незамедлительного безвозмездного устранения недостатков товара;`,
                                    font: "Times New Roman",
                                    size: 28, // 14 pt = 28 half-points
                                }),
                            ],
                        }),
                        new Paragraph({
                            alignment: AlignmentType.JUSTIFIED,
                            indent: { firstLine: 708 },
                            children: [
                                new TextRun({
                                    text: `1.4. возмещения расходов по устранению недостатков товара.\n`,
                                    font: "Times New Roman",
                                    size: 28, // 14 pt = 28 half-points
                                }),
                            ],
                        }),
                        new Paragraph({
                            alignment: AlignmentType.JUSTIFIED,
                            indent: { firstLine: 708 },
                            children: [
                                new TextRun({
                                    text: `Согласно пункту 3 статьи 20 Закона, потребитель вправе расторгнуть договор розничной купли-продажи и потребовать возврата уплаченной за товар денежной суммы в соответствии с пунктом 4 статьи 27 настоящего Закона.`,
                                    font: "Times New Roman",
                                    size: 28, // 14 pt = 28 half-points
                                }),
                            ],
                        }),
                        new Paragraph({
                            alignment: AlignmentType.JUSTIFIED,
                            indent: { firstLine: 708 },
                            children: [
                                new TextRun({
                                    text: `В соответствии с пунктом 11 статьи 20 Закона, продавец (изготовитель, поставщик, представитель) обязан принять товар ненадлежащего качества у потребителя, а в случае необходимости - провести проверку качества товара, в том числе с привлечением ремонтной организации. Продавец (изготовитель, поставщик, представитель) обязан проинформировать потребителя о его праве на участие в проведении проверки качества товара, а если такая проверка не может быть проведена незамедлительно, - также о месте и времени проведения проверки качества товара. Ремонтная организация при получении товара от продавца (изготовителя, поставщика, представителя) для проведения проверки качества товара обязана провести такую проверку в течение трех дней со дня получения товара.`,
                                    font: "Times New Roman",
                                    size: 28, // 14 pt = 28 half-points
                                }),
                            ],
                        }),
                        new Paragraph({
                            alignment: AlignmentType.JUSTIFIED,
                            indent: { firstLine: 708 },
                            children: [
                                new TextRun({
                                    text: `При возникновении между потребителем и продавцом (изготовителем, поставщиком, представителем) спора о наличии недостатков товара и причинах их возникновения продавец (изготовитель, поставщик, представитель) обязан провести экспертизу товара за свой счет в порядке, установленном Правительством Республики Беларусь. О месте и времени проведения экспертизы потребитель должен быть извещен в письменной форме.`,
                                    font: "Times New Roman",
                                    size: 28, // 14 pt = 28 half-points
                                }),
                            ],
                        }),
                        new Paragraph({
                            alignment: AlignmentType.JUSTIFIED,
                            indent: { firstLine: 708 },
                            children: [
                                new TextRun({
                                    text: `Стоимость экспертизы оплачивается продавцом (изготовителем, поставщиком, представителем). Если в результате проведенной экспертизы товара установлено, что недостатки товара отсутствуют или возникли после передачи товара потребителю вследствие нарушения им установленных правил использования, хранения, транспортировки товара или действий третьих лиц либо непреодолимой силы, потребитель обязан возместить продавцу (изготовителю, поставщику, представителю) расходы на проведение экспертизы, а также связанные с ее проведением расходы на транспортировку товара.`,
                                    font: "Times New Roman",
                                    size: 28, // 14 pt = 28 half-points
                                }),
                            ],
                        }),
                        new Paragraph({
                            alignment: AlignmentType.JUSTIFIED,
                            indent: { firstLine: 708 },
                            children: [
                                new TextRun({
                                    text: `Потребитель вправе принять участие в проведении проверки качества и экспертизы товара лично или через своего представителя, оспорить заключение экспертизы товара в судебном порядке, а также при возникновении между потребителем и продавцом (изготовителем, поставщиком, представителем) спора о наличии недостатков товара и причинах их возникновения провести экспертизу товара за свой счет. Если в результате экспертизы товара, проведенной за счет потребителя, установлено, что недостатки возникли до передачи товара потребителю или по причинам, возникшим до момента его передачи, продавец (изготовитель, поставщик, представитель) обязан возместить потребителю расходы на проведение экспертизы, а также связанные с ее проведением расходы на транспортировку товара.`,
                                    font: "Times New Roman",
                                    size: 28, // 14 pt = 28 half-points
                                }),
                            ],
                        }),
                        new Paragraph({
                            alignment: AlignmentType.JUSTIFIED,
                            indent: { firstLine: 708 },
                            children: [
                                new TextRun({
                                    text: `Согласно статье 21 Закона, потребитель вправе предъявить предусмотренные статьей 20 Закона требования продавцу (изготовителю, поставщику, представителю) в отношении недостатков товара в течение гарантийного срока или срока годности товара.`,
                                    font: "Times New Roman",
                                    size: 28, // 14 pt = 28 half-points
                                }),
                            ],
                        }),
                        new Paragraph({
                            alignment: AlignmentType.JUSTIFIED,
                            indent: { firstLine: 708 },
                            children: [
                                new TextRun({
                                    text: `Согласно пункту 2 статьи 15 Закона, убытки, причиненные потребителю, подлежат возмещению в полном объеме сверх неустойки, установленной законодательством или договором.`,
                                    font: "Times New Roman",
                                    size: 28, // 14 pt = 28 half-points
                                }),
                            ],
                            spacing: { after: 300 }
                        }),
                        new Paragraph({
                            alignment: AlignmentType.JUSTIFIED,
                            children: [
                                new TextRun({
                                    text: `ПРОШУ:`,
                                    font: "Times New Roman",
                                    size: 28, // 14 pt = 28 half-points
                                }),
                            ],
                        }),
                        ...paragraphs,
                        new Paragraph({
                            alignment: AlignmentType.JUSTIFIED,
                            indent: { firstLine: 708 },
                            children: [
                                new TextRun({
                                    text: `Одновременно информирую, что в случае неисполнения Вами заявленных требований мною будет подготовлено исковое заявление в суд.\n
                                    
                                    \nВ этом случае дополнительно будут заявлены требования о компенсации морального вреда. 
                                    `,
                                    font: "Times New Roman",
                                    size: 28, // 14 pt = 28 half-points
                                }),
                            ],
                            spacing: { after: 300 }
                        }),
                        new Paragraph({
                            alignment: AlignmentType.JUSTIFIED,
                            indent: { firstLine: 708 },
                            children: [
                                new TextRun({
                                    text: `Срок предоставления ответа на претензию составляет ${daysForOtvet} дней с момента ее получения. `,
                                    font: "Times New Roman",
                                    size: 28, // 14 pt = 28 half-points
                                }),
                            ],
                            spacing: { after: 500 }
                        }),
                        new Table({
                            width: { size: 100, type: WidthType.PERCENTAGE }, // Таблица на всю ширину страницы
                            rows: [
                                new TableRow({
                                    children: [
                                        new TableCell({
                                            children: [
                                                new Paragraph({
                                                    alignment: AlignmentType.LEFT,
                                                    children: [
                                                        new TextRun({
                                                            text: date,
                                                            font: "Times New Roman",
                                                            size: 28, // 14 pt = 28 half-points
                                                        }),
                                                    ],
                                                })
                                            ],
                                            borders: {
                                                top: {style: BorderStyle.NONE, size: 0, color: "FFFFFF"},
                                                bottom: {style: BorderStyle.NONE, size: 0, color: "FFFFFF"},
                                                left: {style: BorderStyle.NONE, size: 0, color: "FFFFFF"},
                                                right: {style: BorderStyle.NONE, size: 0, color: "FFFFFF"},
                                            },
                                            margins: { top: 0, bottom: 0, left: 0, right: 0 },
                                            verticalAlign: "top",
                                            width: { size: 100 / 3, type: WidthType.PERCENTAGE }
                                        }),
                                        new TableCell({
                                            children: [
                                                new Paragraph({
                                                    alignment: AlignmentType.CENTER,
                                                    children: [
                                                        new TextRun({
                                                            text: `РОЗПРОЗП`,
                                                            font: "Times New Roman",
                                                            size: 28,
                                                            color: "FFFFFF" // Белый цвет текста
                                                        }),
                                                    ],
                                                })
                                            ],
                                            borders: {
                                                top: {style: BorderStyle.NONE, size: 0, color: "FFFFFF"},
                                                bottom: {style: BorderStyle.SINGLE, size: 1, color: "000000"},
                                                left: {style: BorderStyle.NONE, size: 0, color: "FFFFFF"},
                                                right: {style: BorderStyle.NONE, size: 0, color: "FFFFFF"},
                                            },
                                            margins: { top: 0, bottom: 0, left: 0, right: 0 },
                                            verticalAlign: "top",
                                            width: { size: 100 / 3, type: WidthType.PERCENTAGE }
                                        }),
                                        new TableCell({
                                            children: [
                                                new Paragraph({
                                                    alignment: AlignmentType.RIGHT,
                                                    children: [
                                                        new TextRun({
                                                            text: abr,
                                                            font: "Times New Roman",
                                                            size: 28, // 14 pt = 28 half-points
                                                        }),
                                                    ],
                                                })
                                            ],
                                            borders: {
                                                top: {style: BorderStyle.NONE, size: 0, color: "FFFFFF"},
                                                bottom: {style: BorderStyle.NONE, size: 0, color: "FFFFFF"},
                                                left: {style: BorderStyle.NONE, size: 0, color: "FFFFFF"},
                                                right: {style: BorderStyle.NONE, size: 0, color: "FFFFFF"},
                                            },
                                            margins: { top: 0, bottom: 0, left: 0, right: 0 },
                                            verticalAlign: "top",
                                            width: { size: 100 / 3, type: WidthType.PERCENTAGE }
                                        }),
                                    ],
                                }),
                            ],
                        }), 
                        new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                                new TextRun({
                                    text: `(подпись)`,
                                    font: "Times New Roman",
                                    size: 20, // 14 pt = 28 half-points
                                    italics: true,
                                    bold: true
                                }),
                            ],
                        }),
                    ],
                    page: {
                        size: { width: 16838, height: 23811 } // A4 size in twips (1 cm = 567 twips)
                    }
                },
            ],
        });

        Packer.toBlob(doc).then((blob) => {
            saveAs(blob, "Претензия.docx");
        });

        
    });
});



function abbreviateName(fullName) {
    // Split the full name into parts
    const parts = fullName.split(' ');

    // Ensure there are exactly 3 parts (last name, first name, middle name)
    if (parts.length !== 3) {
        throw new Error('Пожалуйста введите свои полные ФИО: фамилию, имя, отчество!');
    }

    // Extract last name, first name, and middle name
    const lastName = parts[0];
    const firstNameInitial = parts[1].charAt(0);
    const middleNameInitial = parts[2].charAt(0);

    // Construct the abbreviated name
    const abbreviatedName = `${lastName} ${firstNameInitial}.${middleNameInitial}.`;

    return abbreviatedName;
}

function getCurrentDate() {
    const date = new Date();
    
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0'); // Months are zero-indexed, so add 1
    const year = date.getFullYear();

    return `${day}.${month}.${year}`;
}

function allClear() {
    document.querySelector("#name").value = "";
    document.querySelector("#address").value = "";
    document.querySelector("#phone").value = "";
    document.querySelector("#liabelee").value = "";
    document.querySelector("#liabeleeAddress").value = "";
    document.querySelector("#complaint").value = "";
    document.querySelector("#unp").value = "";
    document.querySelector("#price").value = "";
    document.querySelector("#good").value = "";
    document.querySelector("#date").value = "";
    document.querySelector("#sorazm").value = "";
    document.querySelector("#vozmesh").value = "";
}