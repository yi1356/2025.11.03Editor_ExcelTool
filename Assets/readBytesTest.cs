using System.Collections;
using System.Collections.Generic;
using Table;
using UnityEngine;

public class readBytesTest : MonoBehaviour
{
    // Start is called before the first frame update
    void Start()
    {
        List<weapon> weapons = weapon.LoadBytes();
        for (int i = 0; i < weapons.Count; i++)
        {
            weapon w = weapons[i];
            string log = w.id + "," + w.name + "," + w.prefabName + ",";
            for (int d = 0; d < w.desc.Count; d++)//Z：这里的数组是一行中的desc这一格中的内容
            {
                log += w.desc[d] + ",";
            }
            for (int n = 0; n < w.desc.Count; n++)
            {
                log += w.desc[n] + ",";
            }
            Debug.Log(log);
        }
    }
}
