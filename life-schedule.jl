using JuMP, GLPKMathProgInterface, DataFrames, Taro, ConditionalJuMP

Taro.init()

# get staff availability tables
staff_array = [
    DataFrame(Taro.readxl("availability.xlsx", "availability", "B2:F24", header = false)),
    DataFrame(Taro.readxl("availability.xlsx", "availability", "I2:M24", header = false)),
    DataFrame(Taro.readxl("availability.xlsx", "availability", "P2:T24", header = false)),
    DataFrame(Taro.readxl("availability.xlsx", "availability", "W2:AA24", header = false)),
    DataFrame(Taro.readxl("availability.xlsx", "availability", "AD2:AH24", header = false)),
    DataFrame(Taro.readxl("availability.xlsx", "availability", "AK2:AO24", header = false)),
    DataFrame(Taro.readxl("availability.xlsx", "availability", "AR2:AV24", header = false)),
    DataFrame(Taro.readxl("availability.xlsx", "availability", "AY2:BC24", header = false))]

# number of staff variable, assume working period remains constant
staff = length(staff_array)

# create 3D availability array
av_matrix = Array{Int8}(undef, 23, 5, staff)

for k in 1:staff
    for i in 1:23
        for j in 1:5
            av_matrix[i,j,k] = Matrix(staff_array[k])[i,j]
        end
    end
end

# optimization model
m = Model(solver = GLPKSolverMIP())

# 23 x 5 x staff binary assignment 3d matrix
# 1 if employee k assigned to shift (i,j), 0 otherwise
@variable(m, x[1:23, 1:5, 1:staff], Bin)

# maximize preference score sum
@objective(m, Max, sum(av_matrix[i, j, k] * x[i, j, k] for i in 1:23, j in 1:5, k in 1:staff))

# constraints

# cons1: each person (except Ty = 8) works 10hrs per week
for k in 1:staff-1
    @constraint(m, sum(x[i, j, k] for i in 1:23, j in 1:5) == 20)
end

# cons1.1: Ty = 8 works max 13hrs per week (no min)
@constraint(m, sum(x[i, j, 8] for i in 1:23, j in 1:5) <= 26)

# cons2: 1-3 people working at any given time
for i in 1:23
    for j in 1:5
        @constraint(m, 1 <= sum(x[i, j, k] for k in 1:staff) <= 3)
    end
end

# !!! cons3: each shift is at least 1hr

# !!! cons4: each shift is at most 4hrs

status = solve(m)

println("Objective value: ", getobjectivevalue(m))
assn_matrix = Array{Int64}(getvalue(x))

# create final assignment array

# flatten 3d matrix into 2d array
