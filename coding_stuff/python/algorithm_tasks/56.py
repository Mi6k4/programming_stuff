def function(intervals):
    output_inntervals=[]
    i=0
    for i in range(len(intervals)):
        print(intervals[i])
        #print(intervals[i][-1])
        #print(intervals[i+1][0])
        if intervals[i][-1]>=intervals[i+1][0]:
            output_inntervals.append([intervals[i][0],intervals[i+1][-1]])
        elif intervals[i][0]<intervals[i-1][-1]:
            pass
        else:
            output_inntervals.append(intervals[i])
        print(output_inntervals)
        i+=1
    print(output_inntervals)
    return output_inntervals


def function_2(intervals):
    output_inntervals=[]
    if len(intervals)==0 or len(intervals)==1:
        return intervals
    else:
        for i in range(1,len(intervals)):
            print(i)
            print(intervals[i])
            #print(intervals[i][-1])
            #print(intervals[i+1][0])
            if intervals[i][0]<=intervals[i-1][-1] and intervals[i][0]>intervals[i-1][0] and  intervals[i][-1]>intervals[i-1][-1]:
                output_inntervals.append([intervals[i-1][0],intervals[i][-1]])
            elif intervals[i][0]<intervals[i-1][0] and intervals[i][-1]>=intervals[i-1][-1]:
                output_inntervals.append([intervals[i][0],intervals[i][-1]])
            elif intervals[i] == intervals[i - 1]:
                output_inntervals.append(intervals[i])
            elif intervals[i][0]<intervals[i-1][-1] and intervals[i][0]<intervals[i-1][0] and intervals[i][-1]<intervals[i-1][-1]:
                output_inntervals.append([intervals[i][0],intervals[i-1][-1]])
            elif intervals[i][0]>intervals[i-1][0] and intervals[i][-1]<intervals[i-1][-1]:
                output_inntervals.append([intervals[i-1][0],intervals[i-1][-1]])
            elif  intervals[i][0]<intervals[i-1][0] and intervals[i][-1]>intervals[i-1][-1]:
                output_inntervals.append([intervals[i][0],intervals[i][-1]])
            elif intervals[i][0]>intervals[i-1][-1]:
                if len(output_inntervals)==0 and intervals[i-1][-1]:
                    output_inntervals.append(intervals[i-1])
                else:
                    pass
                output_inntervals.append(intervals[i])
            #else:
                #output_inntervals.append(intervals[i])
            print(output_inntervals)
            i+=1
        print(output_inntervals)
        return output_inntervals

def function_3(intervals):
    #intervals.sort(key = lambda x: x[0])
    #output_intervals = []
    #for i in range(1,len(intervals)):
    #    if intervals[i][0]<=intervals[i-1][-1] and intervals[i][-1]>intervals[i-1][-1]:
    #        output_intervals.append([intervals[i-1][0],intervals[i][-1]])
    #    elif intervals[i][0]<=intervals[i-1][-1] and intervals[i][-1]<intervals[i-1][-1]:
    #        output_intervals.append(intervals[i-1])
    #print(intervals, ' ', output_intervals)
    intervals.sort(key=lambda x: x[0])
    res = [intervals[0]]

    for i in range(1, len(intervals)):
        interval = intervals[i]

        if res[-1][1] < interval[0]:
            res.append(interval)
        else:
            res[-1][1] = max(res[-1][1], interval[1])

    return res


intervals_1 = [[1, 3], [2, 6], [8, 10], [15, 18]]
intervals_2 = [[1,4],[4,5]]
intervals_3 = [[1,4],[5,6]]
intervals_4 =[[1,4],[0,4]]
intervals_5 = [[1,4],[1,4]]
intervals_6 = [[1,4],[0,1]]
intervals_7 = [[1,4],[2,3]]
intervals_8 = [[1,4],[0,5]]

function_2(intervals_1)
function_2(intervals_2)
function_2(intervals_3)
function_2(intervals_4)
function_2(intervals_5)
function_2(intervals_6)
function_2(intervals_7)
function_2(intervals_8)


function_3(intervals_1)
function_3(intervals_2)
function_3(intervals_3)
function_3(intervals_4)
function_3(intervals_5)
function_3(intervals_6)
function_3(intervals_7)
function_3(intervals_8)


