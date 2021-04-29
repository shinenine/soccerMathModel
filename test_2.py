import pandas as pd
import numpy as np

excel_path = 'F:\\study\\数学建模\\2021\\tongjimcm2021D体能数据.xlsx'


def mScore(val):
    s = 0
    if val == 0:
        return 0
    if val <= 3.92:
        s = 10
    elif val <= 4.07:
        s = 9
    elif val <= 4.17:
        s = 8
    elif val <= 4.22:
        s = 7
    elif val <= 4.34:
        s = 6
    elif val <= 4.42:
        s = 5
    elif val <= 4.52:
        s = 4
    elif val <= 4.67:
        s = 3
    elif val <= 4.77:
        s = 2
    elif val <= 4.87:
        s = 1
    else:
        s = 0
    return s


def arrowRunningScore(val):
    s = 0
    if val == 0:
        return 0
    if val <= 7.91:
        s = 10
    elif val <= 8.01:
        s = 9
    elif val <= 8.11:
        s = 8
    elif val <= 8.21:
        s = 7
    elif val <= 8.31:
        s = 6
    elif val <= 8.41:
        s = 5
    elif val <= 8.55:
        s = 4
    elif val <= 8.73:
        s = 3
    elif val <= 8.89:
        s = 2
    elif val <= 9.05:
        s = 1
    else:
        s = 0
    return s


def jumpingScore(val):
    s = 0
    if val == 0:
        return 0
    if val >= 2.81:
        s = 10
    elif val >= 2.74:
        s = 9
    elif val >= 2.66:
        s = 8
    elif val >= 2.59:
        s = 7
    elif val >= 2.49:
        s = 6
    elif val >= 2.43:
        s = 5
    elif val >= 2.38:
        s = 4
    elif val >= 2.32:
        s = 3
    elif val >= 2.26:
        s = 2
    elif val >= 2.21:
        s = 1
    else:
        s = 0
    return s


def upJumpingScore(val):
    s = 0
    if val == 0:
        return 0
    if val >= 65:
        s = 10
    elif val >= 60.1:
        s = 9
    elif val >= 55.1:
        s = 8
    elif val >= 51.1:
        s = 7
    elif val >= 47.1:
        s = 6
    elif val >= 45.6:
        s = 5
    elif val >= 43.8:
        s = 4
    elif val >= 42:
        s = 3
    elif val >= 40.2:
        s = 2
    elif val >= 38.4:
        s = 1
    else:
        s = 0
    return s


def upScore(val):
    s = 0
    if val == 0:
        return 0
    if val >= 19:
        s = 10
    elif val >= 16:
        s = 9
    elif val >= 14:
        s = 8
    elif val >= 11:
        s = 7
    elif val >= 10:
        s = 6
    elif val >= 9:
        s = 5
    elif val >= 8:
        s = 4
    elif val >= 7:
        s = 3
    elif val >= 6:
        s = 2
    elif val >= 5:
        s = 1
    else:
        s = 0
    return s


def yoloScore(val):
    s = 0
    if val == 0:
        return 0
    if val >= 1160:
        s = 10
    elif val >= 1080:
        s = 9
    elif val >= 1000:
        s = 8
    elif val >= 880:
        s = 7
    elif val >= 800:
        s = 6
    elif val >= 720:
        s = 5
    elif val >= 640:
        s = 4
    elif val >= 600:
        s = 3
    elif val >= 520:
        s = 2
    elif val >= 440:
        s = 1
    else:
        s = 0
    return s


def computeTeamScore():
    teams = []

    for i in range(15):
        table = pd.read_excel(excel_path, sheet_name=str(i + 1))
        table = table.fillna(0)
        rows, cols = table.shape
        team = []
        for row in range(rows):
            score = mScore(table.iloc[row, 1]) * 0.15 + arrowRunningScore(table.iloc[row, 2]) * 0.15 + \
                    jumpingScore(table.iloc[row, 3]) * 0.1 + upJumpingScore(table.iloc[row, 4]) * 0.1 + \
                    upScore(table.iloc[row, 5]) * 0.1 + yoloScore(table.iloc[row, 6]) * 0.4
            team.append(score)
        teams.append(team)
    return teams


def writeExcel(teams):
    writer = pd.ExcelWriter('teamScore.xlsx')
    for i in range(len(teams)):
        teamScore = np.array(teams[i])
        data = pd.DataFrame(teamScore)
        data.to_excel(writer, str(i + 1), float_format='%.5f', header=False, index=False)
    writer.save()


def writeFailExcel(teams):
    teamFail = []
    for i in range(len(teams)):
        fail = []
        for k in range(len(teams[i])):
            if teams[i][k] < 6:
                fail.append((k + 1, teams[i][k]))
        teamFail.append(fail)
    failWriter = pd.ExcelWriter('teamFailMember.xlsx')
    for i in range(len(teamFail)):
        teamScore = np.array(teamFail[i])
        data = pd.DataFrame(teamScore)
        data.to_excel(failWriter, str(i + 1), float_format='%.5f', header=False, index=False)
    failWriter.save()


def computePass(teams):
    total = 0
    score = []
    for i in range(len(teams)):
        total += len(teams[i])
        for k in range(len(teams[i])):
            score.append(teams[i][k])
    score.sort()
    passTotal = int(total * 0.15)
    return score[passTotal]


def writePassExcel(teams):
    teamPass = []
    passScore = computePass(teams)
    print(passScore)
    for i in range(len(teams)):
        fail = []
        for k in range(len(teams[i])):
            if teams[i][k] >= passScore:
                fail.append((k + 1, teams[i][k]))
        teamPass.append(fail)
    passWriter = pd.ExcelWriter('teamPassMember_2.xlsx')
    for i in range(len(teamPass)):
        teamScore = np.array(teamPass[i])
        data = pd.DataFrame(teamScore)
        data.to_excel(passWriter, str(i + 1), float_format='%.5f', header=False, index=False)
    passWriter.save()


def writeAverage(teams):
    averageScore = []
    for i in range(len(teams)):
        sumScore = 0
        for k in range(len(teams[i])):
            sumScore += teams[i][k]
        averageScore.append((i + 1, sumScore / len(teams[i])))
    averageWriter = pd.ExcelWriter('teamAverageScore.xlsx')
    teamScore = np.array(averageScore)
    data = pd.DataFrame(teamScore)
    data.to_excel(averageWriter, float_format='%.5f', header=False, index=False)

    averageWriter.save()


if __name__ == '__main__':
    teamsScore = computeTeamScore()
    # writeExcel(teamsScore)
    writePassExcel(teamsScore)
    # writeFailExcel(teamsScore)
    # writeAverage(teamsScore)
