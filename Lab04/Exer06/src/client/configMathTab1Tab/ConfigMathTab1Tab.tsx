import * as React from "react";
import { Provider, Flex, Text, Button, Header, Input, InputProps } from "@fluentui/react-northstar";
import { useState, useEffect } from "react";
import { useTeams } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";

export interface IConfigMathTabState {
  mathOperator ?: string;
  operandA: number;
  operandB: number;
  result: string;
}


/**
 * Implementation of the ConfigMathTab1 content page
 */
export const ConfigMathTab1Tab = () => {

    const [{ inTeams, theme, context }] = useTeams();
    const [entityId, setEntityId] = useState<string | undefined>();
    const [mathTabState, setMathTabState] = useState<IConfigMathTabState>({ mathOperator: "add" } as IConfigMathTabState);

    const handleOnChangedOperandA = (data?: InputProps): void => {
        if (data && !isNaN(Number(data.value))) {
            setMathTabState(state => ({
            ...state,
            operandA: data.value
            } as IConfigMathTabState));
        }
    }

    const handleOnChangedOperandB = (data?: InputProps): void => {
        if (data && !isNaN(Number(data.value))) {
            setMathTabState(state => ({
            ...state,
            operandB: data.value
            } as IConfigMathTabState));
        }
    }

    const handleOperandChange = (): void => {
        let stringResult: string = "n/a";

        if (mathTabState) {
            if (!isNaN(Number(mathTabState.operandA)) && !isNaN(Number(mathTabState.operandB))) {
                switch (mathTabState.mathOperator) {
                    case "add":
                        stringResult = (Number(mathTabState.operandA) + Number(mathTabState.operandB)).toString();
                        break;
                    case "subtract":
                        stringResult = (Number(mathTabState.operandA) - Number(mathTabState.operandB)).toString();
                        break;
                    case "multiply":
                        stringResult = (Number(mathTabState.operandA) * Number(mathTabState.operandB)).toString();
                        break;
                    case "divide":
                        stringResult = (Number(mathTabState.operandA) / Number(mathTabState.operandB)).toString();
                        break;
                    default:
                        stringResult = "n/a";
                        break;
                }
            }
        }
        setMathTabState(state => ({
            ...state,
            result: stringResult
        } as IConfigMathTabState));
    }


    useEffect(() => {
        if (inTeams === true) {
            microsoftTeams.appInitialization.notifySuccess();
        } else {
            setEntityId("Not in Microsoft Teams");
        }
    }, [inTeams]);

    useEffect(() => {
        /*if (context) {
            setEntityId(context.entityId);
        }*/
         if (context) {
            setEntityId(context.entityId);
            setMathTabState(state => ({ 
                ...state, 
                mathOperator: context.entityId.replace("MathPage", "")
            } as IConfigMathTabState));
        }
    }, [context]);

    /**
     * The render() method to create the UI of the tab
     */
    /*return (
        <Provider theme={theme}>
            <Flex fill={true} column styles={{
                padding: ".8rem 0 .8rem .5rem"
            }}>
                <Flex.Item>
                    <Header content="This is your tab" />
                </Flex.Item>
                <Flex.Item>
                    <div>
                        <div>
                            <Text content={entityId} />
                        </div>

                        <div>
                            <Button onClick={() => alert("It worked!")}>A sample button</Button>
                        </div>
                    </div>
                </Flex.Item>
                <Flex.Item styles={{
                    padding: ".8rem 0 .8rem .5rem"
                }}>
                    <Text size="smaller" content="(C) Copyright brotonsmartlife" />
                </Flex.Item>
            </Flex>
        </Provider>
    );*/
    return (
        <Provider theme={theme}>
            <Flex column gap="gap.smaller">
            <Header>This is your tab</Header>
            <Text content="Enter the values to calculate" size="medium"></Text>

            <Flex gap="gap.smaller">
                <Flex.Item>
                <Flex gap="gap.smaller">
                    <Flex.Item>
                    <Input autoFocus
                        value={mathTabState.operandA}
                        onChange={(e, data) => handleOnChangedOperandA(data)}></Input>
                    </Flex.Item>
                    <Text content={mathTabState.mathOperator}></Text>
                    <Flex.Item>
                    <Input value={mathTabState.operandB}
                        onChange={(e, data) => handleOnChangedOperandB(data)}></Input>
                    </Flex.Item>
                </Flex>
                </Flex.Item>
                <Button content="Calculate" primary
                onClick={handleOperandChange}></Button>
                <Text content={mathTabState.result}></Text>
            </Flex>
            <Text content="(C) Copyright BrotonSmartlife" size="smallest"></Text>
            </Flex>
        </Provider>
    );

};
