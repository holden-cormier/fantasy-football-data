import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.w3c.dom.css.Counter

class CreatingFantasyXL {
    static Workbook inputWorkbook = new XSSFWorkbook("data.xlsx")
    static XSSFSheet allDataSheet = inputWorkbook.getSheetAt(0)


    static void main(String[] args) {
        Workbook outputWorkbook = new XSSFWorkbook()
        XSSFSheet outputSheet = outputWorkbook.createSheet()
        XSSFSheet rbOutput = outputWorkbook.createSheet()
        XSSFSheet qbOutput = outputWorkbook.createSheet()
        XSSFSheet wrOutput = outputWorkbook.createSheet()
        XSSFSheet teOutput = outputWorkbook.createSheet()
        outputSheet.createRow(0)
        outputSheet.getRow(0).createCell(0).setCellValue("Ranking")
        outputSheet.getRow(0).createCell(1).setCellValue("Pick")
        outputSheet.getRow(0).createCell(2).setCellValue("Name")
        outputSheet.getRow(0).createCell(3).setCellValue("Pos")
        outputSheet.getRow(0).createCell(4).setCellValue("Team")
        outputSheet.getRow(0).createCell(5).setCellValue("Bye")
        outputSheet.getRow(0).createCell(6).setCellValue("Games Played")
        outputSheet.getRow(0).createCell(7).setCellValue("Pass Comp")
        outputSheet.getRow(0).createCell(8).setCellValue("Pass Att")
        outputSheet.getRow(0).createCell(9).setCellValue("Pass Yard")
        outputSheet.getRow(0).createCell(10).setCellValue("Pass TD")
        outputSheet.getRow(0).createCell(11).setCellValue("INT")
        outputSheet.getRow(0).createCell(12).setCellValue("Rush ATT")
        outputSheet.getRow(0).createCell(13).setCellValue("Rush Yard")
        outputSheet.getRow(0).createCell(14).setCellValue("Rush TD")
        outputSheet.getRow(0).createCell(15).setCellValue("Rec")
        outputSheet.getRow(0).createCell(16).setCellValue("Targets")
        outputSheet.getRow(0).createCell(17).setCellValue("Rec Yards")
        outputSheet.getRow(0).createCell(18).setCellValue("Rec TD")
        outputSheet.getRow(0).createCell(19).setCellValue("Fantasy Points")
        outputSheet.getRow(0).createCell(20).setCellValue("Fantasy PPG")

        int rowCounter = 1
        int rbCounter = 0
        int qbCounter = 0
        int wrCounter = 0
        int teCounter = 0

        int rbPos = 0
        int qbPos = 0
        int wrPos = 0
        int tePos = 0
        XSSFCellStyle style = outputWorkbook.createCellStyle()
        style.setFillForegroundColor(IndexedColors.YELLOW.getIndex())
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND)

        for (Row row : allDataSheet) {
            if (row.getRowNum() > 1) {
                String ranking = row.getCell(0)
                String pick = row.getCell(1)
                String name = row.getCell(2)
                String pos = row.getCell(3)
                String bye = row.getCell(4)
                String team = row.getCell(5)

                PlayerData playerData = new PlayerData(ranking, pick, name, pos, bye, team)
                playerData.loadPlayerStats()


                //put data onto output
                outputSheet.createRow(rowCounter)
                outputSheet.getRow(rowCounter).createCell(0).setCellValue(playerData.getRanking())
                outputSheet.getRow(rowCounter).createCell(1).setCellValue(playerData.getPick())
                outputSheet.getRow(rowCounter).createCell(2).setCellValue(playerData.getName())
                outputSheet.getRow(rowCounter).createCell(3).setCellValue(playerData.getPos())
                outputSheet.getRow(rowCounter).createCell(4).setCellValue(playerData.getTeams()[0])
                outputSheet.getRow(rowCounter).createCell(6).setCellValue(playerData.getGamesPlayed()[0])
                outputSheet.getRow(rowCounter).createCell(7).setCellValue(playerData.getPassComp()[0])
                outputSheet.getRow(rowCounter).createCell(8).setCellValue(playerData.getPassAtt()[0])
                outputSheet.getRow(rowCounter).createCell(9).setCellValue(playerData.getPassYard()[0])
                outputSheet.getRow(rowCounter).createCell(10).setCellValue(playerData.getPassTD()[0])
                outputSheet.getRow(rowCounter).createCell(11).setCellValue(playerData.getPassInt()[0])
                outputSheet.getRow(rowCounter).createCell(12).setCellValue(playerData.getRushAtt()[0])
                outputSheet.getRow(rowCounter).createCell(13).setCellValue(playerData.getRushYard()[0])
                outputSheet.getRow(rowCounter).createCell(14).setCellValue(playerData.getRushTD()[0])
                outputSheet.getRow(rowCounter).createCell(15).setCellValue(playerData.getRec()[0])
                outputSheet.getRow(rowCounter).createCell(16).setCellValue(playerData.getTargets()[0])
                outputSheet.getRow(rowCounter).createCell(17).setCellValue(playerData.getRecYards()[0])
                outputSheet.getRow(rowCounter).createCell(18).setCellValue(playerData.getRecTD()[0])
                outputSheet.getRow(rowCounter).createCell(19).setCellValue(playerData.getFantasyPoints()[0])
                outputSheet.getRow(rowCounter).createCell(20).setCellValue(playerData.getFantasyPPG()[0])
                rowCounter++

                outputSheet.createRow(rowCounter)
                outputSheet.getRow(rowCounter).createCell(4).setCellValue(playerData.getTeams()[1])
                outputSheet.getRow(rowCounter).createCell(6).setCellValue(playerData.getGamesPlayed()[1])
                outputSheet.getRow(rowCounter).createCell(7).setCellValue(playerData.getPassComp()[1])
                outputSheet.getRow(rowCounter).createCell(8).setCellValue(playerData.getPassAtt()[1])
                outputSheet.getRow(rowCounter).createCell(9).setCellValue(playerData.getPassYard()[1])
                outputSheet.getRow(rowCounter).createCell(10).setCellValue(playerData.getPassTD()[1])
                outputSheet.getRow(rowCounter).createCell(11).setCellValue(playerData.getPassInt()[1])
                outputSheet.getRow(rowCounter).createCell(12).setCellValue(playerData.getRushAtt()[1])
                outputSheet.getRow(rowCounter).createCell(13).setCellValue(playerData.getRushYard()[1])
                outputSheet.getRow(rowCounter).createCell(14).setCellValue(playerData.getRushTD()[1])
                outputSheet.getRow(rowCounter).createCell(15).setCellValue(playerData.getRec()[1])
                outputSheet.getRow(rowCounter).createCell(16).setCellValue(playerData.getTargets()[1])
                outputSheet.getRow(rowCounter).createCell(17).setCellValue(playerData.getRecYards()[1])
                outputSheet.getRow(rowCounter).createCell(18).setCellValue(playerData.getRecTD()[1])
                outputSheet.getRow(rowCounter).createCell(19).setCellValue(playerData.getFantasyPoints()[1])
                outputSheet.getRow(rowCounter).createCell(20).setCellValue(playerData.getFantasyPPG()[1])
                rowCounter++

                outputSheet.createRow(rowCounter)
                outputSheet.getRow(rowCounter).createCell(4).setCellValue(playerData.getTeams()[2])
                outputSheet.getRow(rowCounter).createCell(5).setCellValue(playerData.getBye())
                outputSheet.getRow(rowCounter).createCell(6).setCellValue(playerData.getGamesPlayed()[2])
                outputSheet.getRow(rowCounter).createCell(7).setCellValue(playerData.getPassComp()[2])
                outputSheet.getRow(rowCounter).createCell(8).setCellValue(playerData.getPassAtt()[2])
                outputSheet.getRow(rowCounter).createCell(9).setCellValue(playerData.getPassYard()[2])
                outputSheet.getRow(rowCounter).createCell(10).setCellValue(playerData.getPassTD()[2])
                outputSheet.getRow(rowCounter).createCell(11).setCellValue(playerData.getPassInt()[2])
                outputSheet.getRow(rowCounter).createCell(12).setCellValue(playerData.getRushAtt()[2])
                outputSheet.getRow(rowCounter).createCell(13).setCellValue(playerData.getRushYard()[2])
                outputSheet.getRow(rowCounter).createCell(14).setCellValue(playerData.getRushTD()[2])
                outputSheet.getRow(rowCounter).createCell(15).setCellValue(playerData.getRec()[2])
                outputSheet.getRow(rowCounter).createCell(16).setCellValue(playerData.getTargets()[2])
                outputSheet.getRow(rowCounter).createCell(17).setCellValue(playerData.getRecYards()[2])
                outputSheet.getRow(rowCounter).createCell(18).setCellValue(playerData.getRecTD()[2])
                outputSheet.getRow(rowCounter).createCell(19).setCellValue(playerData.getFantasyPoints()[2])
                outputSheet.getRow(rowCounter).createCell(20).setCellValue(playerData.getFantasyPPG()[2])
                rowCounter++


                int counterToUse = 0
                int posToUse = 0
                XSSFSheet sheetToUse
                if(pos == "RB"){
                    sheetToUse = rbOutput
                    counterToUse = rbCounter
                    posToUse = rbPos
                    rbCounter+=3
                    rbPos++
                }
                else if (pos == "QB"){
                    sheetToUse = qbOutput
                    counterToUse = qbCounter
                    posToUse = qbPos
                    qbCounter+=3
                    qbPos++
                }
                else if (pos == "WR"){
                    sheetToUse = wrOutput
                    counterToUse = wrCounter
                    posToUse = wrPos
                    wrCounter+=3
                    wrPos++
                }
                else if (pos == "TE"){
                    sheetToUse = teOutput
                    posToUse = tePos
                    counterToUse = teCounter
                    teCounter+=3
                    tePos++
                }
                if (sheetToUse != null) {
                    sheetToUse.createRow(counterToUse)
                    sheetToUse.getRow(counterToUse).createCell(0).setCellValue(posToUse)
                    sheetToUse.getRow(counterToUse).createCell(1).setCellValue(playerData.getPick())
                    sheetToUse.getRow(counterToUse).createCell(2).setCellValue(playerData.getName())
                    sheetToUse.getRow(counterToUse).createCell(3).setCellValue(playerData.getPos())
                    sheetToUse.getRow(counterToUse).createCell(4).setCellValue(playerData.getTeams()[0])
                    sheetToUse.getRow(counterToUse).createCell(6).setCellValue(playerData.getGamesPlayed()[0])
                    sheetToUse.getRow(counterToUse).createCell(7).setCellValue(playerData.getPassComp()[0])
                    sheetToUse.getRow(counterToUse).createCell(8).setCellValue(playerData.getPassAtt()[0])
                    sheetToUse.getRow(counterToUse).createCell(9).setCellValue(playerData.getPassYard()[0])
                    sheetToUse.getRow(counterToUse).createCell(10).setCellValue(playerData.getPassTD()[0])
                    sheetToUse.getRow(counterToUse).createCell(11).setCellValue(playerData.getPassInt()[0])
                    sheetToUse.getRow(counterToUse).createCell(12).setCellValue(playerData.getRushAtt()[0])
                    sheetToUse.getRow(counterToUse).createCell(13).setCellValue(playerData.getRushYard()[0])
                    sheetToUse.getRow(counterToUse).createCell(14).setCellValue(playerData.getRushTD()[0])
                    sheetToUse.getRow(counterToUse).createCell(15).setCellValue(playerData.getRec()[0])
                    sheetToUse.getRow(counterToUse).createCell(16).setCellValue(playerData.getTargets()[0])
                    sheetToUse.getRow(counterToUse).createCell(17).setCellValue(playerData.getRecYards()[0])
                    sheetToUse.getRow(counterToUse).createCell(18).setCellValue(playerData.getRecTD()[0])
                    sheetToUse.getRow(counterToUse).createCell(19).setCellValue(playerData.getFantasyPoints()[0])
                    sheetToUse.getRow(counterToUse).createCell(20).setCellValue(playerData.getFantasyPPG()[0])
                    counterToUse++

                    sheetToUse.createRow(counterToUse)
                    sheetToUse.getRow(counterToUse).createCell(4).setCellValue(playerData.getTeams()[1])
                    sheetToUse.getRow(counterToUse).createCell(6).setCellValue(playerData.getGamesPlayed()[1])
                    sheetToUse.getRow(counterToUse).createCell(7).setCellValue(playerData.getPassComp()[1])
                    sheetToUse.getRow(counterToUse).createCell(8).setCellValue(playerData.getPassAtt()[1])
                    sheetToUse.getRow(counterToUse).createCell(9).setCellValue(playerData.getPassYard()[1])
                    sheetToUse.getRow(counterToUse).createCell(10).setCellValue(playerData.getPassTD()[1])
                    sheetToUse.getRow(counterToUse).createCell(11).setCellValue(playerData.getPassInt()[1])
                    sheetToUse.getRow(counterToUse).createCell(12).setCellValue(playerData.getRushAtt()[1])
                    sheetToUse.getRow(counterToUse).createCell(13).setCellValue(playerData.getRushYard()[1])
                    sheetToUse.getRow(counterToUse).createCell(14).setCellValue(playerData.getRushTD()[1])
                    sheetToUse.getRow(counterToUse).createCell(15).setCellValue(playerData.getRec()[1])
                    sheetToUse.getRow(counterToUse).createCell(16).setCellValue(playerData.getTargets()[1])
                    sheetToUse.getRow(counterToUse).createCell(17).setCellValue(playerData.getRecYards()[1])
                    sheetToUse.getRow(counterToUse).createCell(18).setCellValue(playerData.getRecTD()[1])
                    sheetToUse.getRow(counterToUse).createCell(19).setCellValue(playerData.getFantasyPoints()[1])
                    sheetToUse.getRow(counterToUse).createCell(20).setCellValue(playerData.getFantasyPPG()[1])
                    counterToUse++

                    sheetToUse.createRow(counterToUse)
                    sheetToUse.getRow(counterToUse).createCell(4).setCellValue(playerData.getTeams()[2])
                    sheetToUse.getRow(counterToUse).createCell(5).setCellValue(playerData.getBye())
                    sheetToUse.getRow(counterToUse).createCell(6).setCellValue(playerData.getGamesPlayed()[2])
                    sheetToUse.getRow(counterToUse).createCell(7).setCellValue(playerData.getPassComp()[2])
                    sheetToUse.getRow(counterToUse).createCell(8).setCellValue(playerData.getPassAtt()[2])
                    sheetToUse.getRow(counterToUse).createCell(9).setCellValue(playerData.getPassYard()[2])
                    sheetToUse.getRow(counterToUse).createCell(10).setCellValue(playerData.getPassTD()[2])
                    sheetToUse.getRow(counterToUse).createCell(11).setCellValue(playerData.getPassInt()[2])
                    sheetToUse.getRow(counterToUse).createCell(12).setCellValue(playerData.getRushAtt()[2])
                    sheetToUse.getRow(counterToUse).createCell(13).setCellValue(playerData.getRushYard()[2])
                    sheetToUse.getRow(counterToUse).createCell(14).setCellValue(playerData.getRushTD()[2])
                    sheetToUse.getRow(counterToUse).createCell(15).setCellValue(playerData.getRec()[2])
                    sheetToUse.getRow(counterToUse).createCell(16).setCellValue(playerData.getTargets()[2])
                    sheetToUse.getRow(counterToUse).createCell(17).setCellValue(playerData.getRecYards()[2])
                    sheetToUse.getRow(counterToUse).createCell(18).setCellValue(playerData.getRecTD()[2])
                    sheetToUse.getRow(counterToUse).createCell(19).setCellValue(playerData.getFantasyPoints()[2])
                    sheetToUse.getRow(counterToUse).createCell(20).setCellValue(playerData.getFantasyPPG()[2])
                    counterToUse++
                }
            }
        }


        FileOutputStream output = new FileOutputStream("./output.xlsx")
        outputWorkbook.write(output)

    }
}

class PlayerData {
    static Workbook inputWorkbook = new XSSFWorkbook("data.xlsx")
    static XSSFSheet allDataSheet = inputWorkbook.getSheetAt(0)
    static XSSFSheet qb2021 = inputWorkbook.getSheetAt(1)
    static XSSFSheet rb2021 = inputWorkbook.getSheetAt(2)
    static XSSFSheet wr2021 = inputWorkbook.getSheetAt(3)
    static XSSFSheet te2021 = inputWorkbook.getSheetAt(4)
    static XSSFSheet k2021 = inputWorkbook.getSheetAt(5)
    static XSSFSheet d2021 = inputWorkbook.getSheetAt(6)
    static XSSFSheet qb2020 = inputWorkbook.getSheetAt(7)
    static XSSFSheet rb2020 = inputWorkbook.getSheetAt(8)
    static XSSFSheet wr2020 = inputWorkbook.getSheetAt(9)
    static XSSFSheet te2020 = inputWorkbook.getSheetAt(10)
    static XSSFSheet k2020 = inputWorkbook.getSheetAt(11)
    static XSSFSheet d2020 = inputWorkbook.getSheetAt(12)
    static XSSFSheet qb2022 = inputWorkbook.getSheetAt(13)
    static XSSFSheet rb2022 = inputWorkbook.getSheetAt(14)
    static XSSFSheet wr2022 = inputWorkbook.getSheetAt(15)
    static XSSFSheet te2022 = inputWorkbook.getSheetAt(16)
    static XSSFSheet k2022 = inputWorkbook.getSheetAt(17)
    static XSSFSheet d2022 = inputWorkbook.getSheetAt(18)
    String ranking
    String pick
    String name
    String pos
    String bye
    String team
    ArrayList<String> teams = new ArrayList<>(["-", "-", "-"])
    ArrayList<String> gamesPlayed = new ArrayList<String>(["-", "-", "-"])
    ArrayList<String> gamesPlayedDiff = new ArrayList<String>(["-", "-", "-"])
    ArrayList<String> passComp = new ArrayList<String>(["-", "-", "-"])
    ArrayList<String> passCompDiff = new ArrayList<String>(["-", "-", "-"])
    ArrayList<String> passAtt = new ArrayList<String>(["-", "-", "-"])
    ArrayList<String> passAttDiff = new ArrayList<String>(["-", "-", "-"])
    ArrayList<String> passYard = new ArrayList<String>(["-", "-", "-"])
    ArrayList<String> passYardDiff = new ArrayList<String>(["-", "-", "-"])
    ArrayList<String> passTD = new ArrayList<String>(["-", "-", "-"])
    ArrayList<String> passTDDiff = new ArrayList<String>(["-", "-", "-"])
    ArrayList<String> passInt = new ArrayList<String>(["-", "-", "-"])
    ArrayList<String> passIntDiff = new ArrayList<String>(["-", "-", "-"])
    ArrayList<String> rushAtt = new ArrayList<String>(["-", "-", "-"])
    ArrayList<String> rushAttDiff = new ArrayList<String>(["-", "-", "-"])
    ArrayList<String> rushYard = new ArrayList<String>(["-", "-", "-"])
    ArrayList<String> rushYardDiff = new ArrayList<String>(["-", "-", "-"])
    ArrayList<String> rushTD = new ArrayList<String>(["-", "-", "-"])
    ArrayList<String> rushTDDiff = new ArrayList<String>(["-", "-", "-"])
    ArrayList<String> rec = new ArrayList<String>(["-", "-", "-"])
    ArrayList<String> recDiff = new ArrayList<String>(["-", "-", "-"])
    ArrayList<String> targets = new ArrayList<String>(["-", "-", "-"])
    ArrayList<String> targetsDiff = new ArrayList<String>(["-", "-", "-"])
    ArrayList<String> recYards = new ArrayList<String>(["-", "-", "-"])
    ArrayList<String> recYardDiff = new ArrayList<String>(["-", "-", "-"])
    ArrayList<String> recTD = new ArrayList<String>(["-", "-", "-"])
    ArrayList<String> recTDDiff = new ArrayList<String>(["-", "-", "-"])
    ArrayList<String> totalTD = new ArrayList<String>(["-", "-", "-"])
    ArrayList<String> totalTDDiff = new ArrayList<String>(["-", "-", "-"])
    ArrayList<String> fantasyPoints = new ArrayList<String>(["-", "-", "-"])
    ArrayList<String> fantasyPointsDiff = new ArrayList<String>(["-", "-", "-"])
    ArrayList<String> fantasyPPG = new ArrayList<String>(["-", "-", "-"])
    ArrayList<String> fantasyPPGDiff = new ArrayList<String>(["-", "-", "-"])

    PlayerData(String ranking, String pick, String name, String pos, String bye, String team) {
        this.ranking = ranking
        this.pick = pick
        this.name = name
        this.pos = pos
        this.bye = bye
        this.team = team
    }


    void loadPlayerStats() {
        if (this.pos == "RB") {
            int playerIndex2021 = findPLayerIndex(rb2021, this.name)
            int playerIndex2020 = findPLayerIndex(rb2020, this.name)
            int playerIndex2022 = findPLayerIndex(rb2022, this.name)
            if (playerIndex2020 != -1) {
                teams.set(0,rb2020.getRow(playerIndex2020).getCell(1) as String)
                gamesPlayed.set(0, rb2020.getRow(playerIndex2020).getCell(2) as String)
                rushAtt.set(0, rb2020.getRow(playerIndex2020).getCell(3) as String)
                rushYard.set(0, rb2020.getRow(playerIndex2020).getCell(4) as String)
                rushTD.set(0, rb2020.getRow(playerIndex2020).getCell(5) as String)
                targets.set(0, rb2020.getRow(playerIndex2020).getCell(6) as String)
                rec.set(0, rb2020.getRow(playerIndex2020).getCell(7) as String)
                recYards.set(0, rb2020.getRow(playerIndex2020).getCell(8) as String)
                recTD.set(0, rb2020.getRow(playerIndex2020).getCell(9) as String)
                fantasyPoints.set(0, rb2020.getRow(playerIndex2020).getCell(10) as String)
                fantasyPPG.set(0, rb2020.getRow(playerIndex2020).getCell(11) as String)
            }
            if (playerIndex2021 != -1) {
                teams.set(1,rb2021.getRow(playerIndex2021).getCell(1) as String)
                gamesPlayed.set(1, rb2021.getRow(playerIndex2021).getCell(2) as String)
                rushAtt.set(1, rb2021.getRow(playerIndex2021).getCell(3) as String)
                rushYard.set(1, rb2021.getRow(playerIndex2021).getCell(4) as String)
                rushTD.set(1, rb2021.getRow(playerIndex2021).getCell(5) as String)
                targets.set(1, rb2021.getRow(playerIndex2021).getCell(6) as String)
                rec.set(1, rb2021.getRow(playerIndex2021).getCell(7) as String)
                recYards.set(1, rb2021.getRow(playerIndex2021).getCell(8) as String)
                recTD.set(1, rb2021.getRow(playerIndex2021).getCell(9) as String)
                fantasyPoints.set(1, rb2021.getRow(playerIndex2021).getCell(10) as String)
                fantasyPPG.set(1, rb2021.getRow(playerIndex2021).getCell(11) as String)
            }
            if (playerIndex2022 != -1) {
                teams.set(2, team)
                gamesPlayed.set(2, rb2022.getRow(playerIndex2022).getCell(1) as String)
                rushAtt.set(2, rb2022.getRow(playerIndex2022).getCell(2) as String)
                rushYard.set(2, rb2022.getRow(playerIndex2022).getCell(3) as String)
                rushTD.set(2, rb2022.getRow(playerIndex2022).getCell(5) as String)
                targets.set(2, rb2022.getRow(playerIndex2022).getCell(6) as String)
                rec.set(2, rb2022.getRow(playerIndex2022).getCell(7) as String)
                recYards.set(2, rb2022.getRow(playerIndex2022).getCell(8) as String)
                recTD.set(2, rb2022.getRow(playerIndex2022).getCell(11) as String)
                fantasyPoints.set(2, rb2022.getRow(playerIndex2022).getCell(13) as String)
                fantasyPPG.set(2, rb2022.getRow(playerIndex2022).getCell(14) as String)
            }
        }
        else if (this.pos == "QB"){
            int playerIndex2021 = findPLayerIndex(qb2021, this.name)
            int playerIndex2020 = findPLayerIndex(qb2020, this.name)
            int playerIndex2022 = findPLayerIndex(qb2022, this.name)
            if (playerIndex2020 != -1) {
                teams.set(0, qb2021.getRow(playerIndex2020).getCell(1) as String)
                gamesPlayed.set(0, qb2020.getRow(playerIndex2020).getCell(2) as String)
                passComp.set(0, qb2020.getRow(playerIndex2020).getCell(3) as String)
                passAtt.set(0, qb2020.getRow(playerIndex2020).getCell(4) as String)
                passYard.set(0, qb2020.getRow(playerIndex2020).getCell(5) as String)
                passTD.set(0, qb2020.getRow(playerIndex2020).getCell(6) as String)
                passInt.set(0, qb2020.getRow(playerIndex2020).getCell(7) as String)
                rushAtt.set(0, qb2020.getRow(playerIndex2020).getCell(8) as String)
                rushYard.set(0, qb2020.getRow(playerIndex2020).getCell(9) as String)
                rushTD.set(0, qb2020.getRow(playerIndex2020).getCell(10) as String)
                fantasyPoints.set(0, qb2020.getRow(playerIndex2020).getCell(11) as String)
                fantasyPPG.set(0, qb2020.getRow(playerIndex2020).getCell(12) as String)
            }
            if (playerIndex2021 != -1) {
                teams.set(1,qb2021.getRow(playerIndex2021).getCell(1) as String)
                gamesPlayed.set(1, qb2021.getRow(playerIndex2021).getCell(2) as String)
                passComp.set(1, qb2021.getRow(playerIndex2021).getCell(3) as String)
                passAtt.set(1, qb2021.getRow(playerIndex2021).getCell(4) as String)
                passYard.set(1, qb2021.getRow(playerIndex2021).getCell(5) as String)
                passTD.set(1, qb2021.getRow(playerIndex2021).getCell(6) as String)
                passInt.set(1, qb2021.getRow(playerIndex2021).getCell(7) as String)
                rushAtt.set(1, qb2021.getRow(playerIndex2021).getCell(8) as String)
                rushYard.set(1, qb2021.getRow(playerIndex2021).getCell(9) as String)
                rushTD.set(1, qb2021.getRow(playerIndex2021).getCell(10) as String)
                fantasyPoints.set(1, qb2021.getRow(playerIndex2021).getCell(11) as String)
                fantasyPPG.set(1, qb2021.getRow(playerIndex2021).getCell(12) as String)

            }
            if (playerIndex2022 != -1) {
                teams.set(2, team)
                gamesPlayed.set(2, qb2022.getRow(playerIndex2022).getCell(1) as String)
                passComp.set(2, qb2022.getRow(playerIndex2022).getCell(3) as String)
                passAtt.set(2, qb2022.getRow(playerIndex2022).getCell(2) as String)
                passYard.set(2, qb2022.getRow(playerIndex2022).getCell(4) as String)
                passTD.set(2, qb2022.getRow(playerIndex2022).getCell(6) as String)
                passInt.set(2, qb2022.getRow(playerIndex2022).getCell(7) as String)
                rushAtt.set(2, qb2022.getRow(playerIndex2022).getCell(9) as String)
                rushYard.set(2, qb2022.getRow(playerIndex2022).getCell(10) as String)
                rushTD.set(2, qb2022.getRow(playerIndex2022).getCell(12) as String)
                fantasyPoints.set(2, qb2022.getRow(playerIndex2022).getCell(14) as String)
                fantasyPPG.set(2, qb2022.getRow(playerIndex2022).getCell(15) as String)

            }

        }
        else if (this.pos == "WR"){
            int playerIndex2021 = findPLayerIndex(wr2021, this.name)
            int playerIndex2020 = findPLayerIndex(wr2020, this.name)
            int playerIndex2022 = findPLayerIndex(wr2022, this.name)
            if (playerIndex2020 != -1) {
                teams.set(0,wr2020.getRow(playerIndex2020).getCell(1) as String)
                gamesPlayed.set(0, wr2020.getRow(playerIndex2020).getCell(2) as String)
                targets.set(0, wr2020.getRow(playerIndex2020).getCell(3) as String)
                rec.set(0, wr2020.getRow(playerIndex2020).getCell(4) as String)
                recYards.set(0, wr2020.getRow(playerIndex2020).getCell(5) as String)
                recTD.set(0, wr2020.getRow(playerIndex2020).getCell(6) as String)
                rushAtt.set(0, wr2020.getRow(playerIndex2020).getCell(7) as String)
                rushYard.set(0, wr2020.getRow(playerIndex2020).getCell(8) as String)
                rushTD.set(0, wr2020.getRow(playerIndex2020).getCell(9) as String)
                fantasyPoints.set(0, wr2020.getRow(playerIndex2020).getCell(10) as String)
                fantasyPPG.set(0, wr2020.getRow(playerIndex2020).getCell(11) as String)
            }
            if (playerIndex2021 != -1) {
                teams.set(1,wr2021.getRow(playerIndex2021).getCell(1) as String)
                gamesPlayed.set(1, wr2021.getRow(playerIndex2021).getCell(2) as String)
                targets.set(1, wr2021.getRow(playerIndex2021).getCell(3) as String)
                rec.set(1, wr2021.getRow(playerIndex2021).getCell(4) as String)
                recYards.set(1, wr2021.getRow(playerIndex2021).getCell(5) as String)
                recTD.set(1, wr2021.getRow(playerIndex2021).getCell(6) as String)
                rushAtt.set(1, wr2021.getRow(playerIndex2021).getCell(7) as String)
                rushYard.set(1, wr2021.getRow(playerIndex2021).getCell(8) as String)
                rushTD.set(1, wr2021.getRow(playerIndex2021).getCell(9) as String)
                fantasyPoints.set(1, wr2021.getRow(playerIndex2021).getCell(10) as String)
                fantasyPPG.set(1, wr2021.getRow(playerIndex2021).getCell(11) as String)
            }
            if (playerIndex2022 != -1) {
                teams.set(2, team)
                gamesPlayed.set(2, wr2022.getRow(playerIndex2022).getCell(1) as String)
                targets.set(2, wr2022.getRow(playerIndex2022).getCell(2) as String)
                rec.set(2, wr2022.getRow(playerIndex2022).getCell(3) as String)
                recYards.set(2, wr2022.getRow(playerIndex2022).getCell(4) as String)
                recTD.set(2, wr2022.getRow(playerIndex2022).getCell(7) as String)
                rushAtt.set(2, wr2022.getRow(playerIndex2022).getCell(8) as String)
                rushYard.set(2, wr2022.getRow(playerIndex2022).getCell(9) as String)
                rushTD.set(2, wr2022.getRow(playerIndex2022).getCell(11) as String)
                fantasyPoints.set(2, wr2022.getRow(playerIndex2022).getCell(13) as String)
                fantasyPPG.set(2, wr2022.getRow(playerIndex2022).getCell(14) as String)

            }
        }
        else if (this.pos == "TE"){
            int playerIndex2021 = findPLayerIndex(te2021, this.name)
            int playerIndex2020 = findPLayerIndex(te2020, this.name)
            int playerIndex2022 = findPLayerIndex(te2022, this.name)
            if (playerIndex2020 != -1) {
                teams.set(0,te2020.getRow(playerIndex2020).getCell(1) as String)
                gamesPlayed.set(0, te2020.getRow(playerIndex2020).getCell(2) as String)
                targets.set(0, te2020.getRow(playerIndex2020).getCell(3) as String)
                rec.set(0, te2020.getRow(playerIndex2020).getCell(4) as String)
                recYards.set(0, te2020.getRow(playerIndex2020).getCell(5) as String)
                recTD.set(0, te2020.getRow(playerIndex2020).getCell(6) as String)
                fantasyPoints.set(0, te2020.getRow(playerIndex2020).getCell(7) as String)
                fantasyPPG.set(0, te2020.getRow(playerIndex2020).getCell(8) as String)
            }
            if (playerIndex2021 != -1) {
                gamesPlayed.set(1, te2021.getRow(playerIndex2021).getCell(2) as String)
                targets.set(1, te2021.getRow(playerIndex2021).getCell(3) as String)
                rec.set(1, te2021.getRow(playerIndex2021).getCell(4) as String)
                recYards.set(1, te2021.getRow(playerIndex2021).getCell(5) as String)
                recTD.set(1, te2021.getRow(playerIndex2021).getCell(6) as String)
                fantasyPoints.set(1, te2021.getRow(playerIndex2021).getCell(7) as String)
                fantasyPPG.set(1, te2021.getRow(playerIndex2021).getCell(8) as String)

            }
            if (playerIndex2022 != -1) {
                teams.set(2, team)
                gamesPlayed.set(2, te2022.getRow(playerIndex2022).getCell(1) as String)
                targets.set(2, te2022.getRow(playerIndex2022).getCell(2) as String)
                rec.set(2, te2022.getRow(playerIndex2022).getCell(3) as String)
                recYards.set(2, te2022.getRow(playerIndex2022).getCell(4) as String)
                recTD.set(2, te2022.getRow(playerIndex2022).getCell(7) as String)
                fantasyPoints.set(2, te2022.getRow(playerIndex2022).getCell(9) as String)
                fantasyPPG.set(2, te2022.getRow(playerIndex2022).getCell(10) as String)

            }
        }

    }

    static int findPLayerIndex(XSSFSheet playerStats, String playerName) {
        for (Row row : playerStats) {
            for (Cell cell : row) {
                if (cell.getColumnIndex() == 0 && row.getRowNum() > 0) {
                    String testVal = cell.getStringCellValue()
                    String lastName = playerName.split(" ")[1]
                    if (testVal.contains(lastName)) {
                        return row.getRowNum()
                    }
                }
            }
        }
        return -1
    }


    String getRanking() {
        return ranking
    }

    String getPick() {
        return pick
    }

    String getName() {
        return name
    }

    String getPos() {
        return pos
    }

    String getTeam() {
        return team
    }

    String getBye() {
        return bye
    }

    ArrayList<String> getGamesPlayed() {
        return gamesPlayed
    }

    ArrayList<String> getGamesPlayedDiff() {
        return gamesPlayedDiff
    }

    ArrayList<String> getPassComp() {
        return passComp
    }

    ArrayList<String> getPassCompDiff() {
        return passCompDiff
    }

    ArrayList<String> getPassAtt() {
        return passAtt
    }

    ArrayList<String> getPassAttDiff() {
        return passAttDiff
    }

    ArrayList<String> getPassYard() {
        return passYard
    }

    ArrayList<String> getPassYardDiff() {
        return passYardDiff
    }

    ArrayList<String> getPassTD() {
        return passTD
    }

    ArrayList<String> getPassTDDiff() {
        return passTDDiff
    }

    ArrayList<String> getPassInt() {
        return passInt
    }

    ArrayList<String> getPassIntDiff() {
        return passIntDiff
    }

    ArrayList<String> getRushAtt() {
        return rushAtt
    }

    ArrayList<String> getRushAttDiff() {
        return rushAttDiff
    }

    ArrayList<String> getRushYard() {
        return rushYard
    }

    ArrayList<String> getRushYardDiff() {
        return rushYardDiff
    }

    ArrayList<String> getRushTD() {
        return rushTD
    }

    ArrayList<String> getRushTDDiff() {
        return rushTDDiff
    }

    ArrayList<String> getRec() {
        return rec
    }

    ArrayList<String> getRecDiff() {
        return recDiff
    }

    ArrayList<String> getTargets() {
        return targets
    }

    ArrayList<String> getTargetsDiff() {
        return targetsDiff
    }

    ArrayList<String> getRecYards() {
        return recYards
    }

    ArrayList<String> getRecYardDiff() {
        return recYardDiff
    }

    ArrayList<String> getRecTD() {
        return recTD
    }

    ArrayList<String> getRecTDDiff() {
        return recTDDiff
    }

    ArrayList<String> getTotalTD() {
        return totalTD
    }

    ArrayList<String> getTotalTDDiff() {
        return totalTDDiff
    }

    ArrayList<String> getFantasyPoints() {
        return fantasyPoints
    }

    ArrayList<String> getFantasyPointsDiff() {
        return fantasyPointsDiff
    }

    ArrayList<String> getFantasyPPG() {
        return fantasyPPG
    }

    ArrayList<String> getFantasyPPGDiff() {
        return fantasyPPGDiff
    }
}
