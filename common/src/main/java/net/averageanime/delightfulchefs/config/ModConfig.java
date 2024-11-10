package net.averageanime.delightfulchefs.config;

import net.averageanime.delightfulchefs.config.annotations.Description;

public class ModConfig implements Config {

    @Override
    public String getName() {
        return "delightful-chefs-config";
    }

    @Override
    public String getExtension() {
        return "json5";
    }

    @Override
    public String getDirectory() {
        return "delightfulchefs";
    }

}


