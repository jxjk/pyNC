
        %
        O1900(TOOL INFO)
        (DATE 2016.7.5)
        (TOOL INFO)
        #1=171(TOOL NO. STRAT)
        #2=185(TOOL NO. END)
        (LOOP)
        WH[#1LE#2]DO1
            #152=#[#1](SET TOOL H)
            M98P1300(COUNDITION)

            IF[#[800+#151]EQ#152]]GOTO9(如果刀具正在使用就跳过)

            T#151(换刀)
            M6
            #3006=#152(CHANGE TOOL)
            #[800+#151]=#152(设置刀具记忆)
            N9
            #1=#1+1
        END1

        N99
        M99
        %

        